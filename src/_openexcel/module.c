#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <datetime.h>
#include <structmember.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

#include "workbook.h"
#include "worksheet.h"
#include "cell.h"
#include "styles.h"
#include "string_table.h"
#include "reader/reader.h"
#include "writer/writer.h"

/* ========== Forward declarations ========== */
static PyTypeObject PyWorkbookType;
static PyTypeObject PyWorksheetType;
static PyTypeObject PyRowIteratorType;
static PyTypeObject PyXlCellType;

/* ========== PyWorkbookObject ========== */

typedef struct {
    PyObject_HEAD
    OxlWorkbook *wb;
} PyWorkbookObject;

/* ========== PyWorksheetObject ========== */

typedef struct {
    PyObject_HEAD
    PyObject    *owner;  /* ref to PyWorkbookObject to keep C data alive */
    OxlWorksheet *ws;
    OxlWorkbook  *wb;
} PyWorksheetObject;

/* ========== PyRowIteratorObject ========== */

typedef struct {
    PyObject_HEAD
    PyObject     *owner_ws;  /* ref to PyWorksheetObject */
    OxlWorksheet *ws;
    OxlWorkbook  *wb;
    uint32_t      cell_pos;
    uint32_t      cur_row;
    uint32_t      min_row;
    uint32_t      max_row;
    uint16_t      min_col;
    uint16_t      max_col;
} PyRowIteratorObject;

/* ========== PyXlCellObject ========== */

typedef struct {
    PyObject_HEAD
    PyObject    *owner_ws;  /* ref to PyWorksheetObject to keep C data alive */
    OxlWorksheet *ws;
    OxlWorkbook  *wb;
    uint32_t      row;     /* 0-based */
    uint16_t      col;     /* 0-based */
} PyXlCellObject;

/* ========== Column letter utilities ========== */

/* "A" -> 1, "Z" -> 26, "AA" -> 27, "AZ" -> 52, "AAA" -> 703
   Input: uppercase column string. Returns 0 on error. */
static int col_str_to_idx(const char *s) {
    if (!s || !*s) return 0;
    int result = 0;
    for (; *s; s++) {
        if (*s < 'A' || *s > 'Z') return 0;
        result = result * 26 + (*s - 'A' + 1);
    }
    return result;
}

/* 1 -> "A", 26 -> "Z", 27 -> "AA". buf must be >= 5 bytes. Null-terminates. */
static void col_idx_to_str(int idx, char *buf) {
    char tmp[8];
    int len = 0;
    while (idx > 0) {
        int rem = (idx - 1) % 26;
        tmp[len++] = (char)('A' + rem);
        idx = (idx - 1) / 26;
    }
    /* reverse */
    for (int i = 0; i < len; i++)
        buf[i] = tmp[len - 1 - i];
    buf[len] = '\0';
}

/* Parse "A1" or "AA10" into 0-based row and col. Returns 0 on success. */
static int parse_a1_ref(const char *ref, uint32_t *row, uint16_t *col) {
    if (!ref || !*ref) return -1;
    /* Read letters */
    const char *p = ref;
    while (*p && *p >= 'A' && *p <= 'Z') p++;
    if (p == ref) return -1;  /* no letters */
    if (!*p) return -1;       /* no digits */

    /* col string */
    char col_str[8];
    size_t col_len = (size_t)(p - ref);
    if (col_len >= sizeof(col_str)) return -1;
    memcpy(col_str, ref, col_len);
    col_str[col_len] = '\0';

    int col_idx = col_str_to_idx(col_str);
    if (col_idx <= 0) return -1;

    /* row number */
    char *end;
    long row_num = strtol(p, &end, 10);
    if (end == p || *end != '\0' || row_num <= 0) return -1;

    *row = (uint32_t)(row_num - 1);  /* 0-based */
    *col = (uint16_t)(col_idx - 1);  /* 0-based */
    return 0;
}

/* ========== Helper: OxlCell → Python object ========== */

static PyObject *cell_to_python(const OxlCell *c, OxlWorkbook *wb) {
    if (c->formula) {
        size_t flen = strlen(c->formula);
        char *tmp = malloc(flen + 2);
        if (!tmp) return PyErr_NoMemory();
        tmp[0] = '=';
        memcpy(tmp + 1, c->formula, flen + 1);
        PyObject *ret = PyUnicode_FromString(tmp);
        free(tmp);
        return ret;
    }
    switch (c->type) {
    case OXL_CELL_EMPTY:
        Py_RETURN_NONE;
    case OXL_CELL_FLOAT:
        return PyFloat_FromDouble(c->v.f);
    case OXL_CELL_STRING: {
        const char *s = oxl_sst_get(&wb->sst, c->v.s_idx);
        if (!s) Py_RETURN_NONE;
        return PyUnicode_DecodeUTF8(s, (Py_ssize_t)strlen(s), "replace");
    }
    case OXL_CELL_INLINE_STR:
        if (!c->v.s_inline) Py_RETURN_NONE;
        return PyUnicode_DecodeUTF8(c->v.s_inline,
                                    (Py_ssize_t)strlen(c->v.s_inline), "replace");
    case OXL_CELL_BOOL:
        return PyBool_FromLong(c->v.b);
    case OXL_CELL_DATE: {
        const OxlDate *d = &c->v.dt;
        if (d->hour == 0 && d->min == 0 && d->sec == 0 && d->usec == 0) {
            return PyDate_FromDate(d->year, d->month ? d->month : 1,
                                   d->day   ? d->day   : 1);
        }
        return PyDateTime_FromDateAndTime(
            d->year, d->month ? d->month : 1, d->day ? d->day : 1,
            d->hour, d->min, d->sec, (int)d->usec);
    }
    case OXL_CELL_ERROR:
        if (c->v.s_inline)
            return PyUnicode_DecodeUTF8(c->v.s_inline,
                                        (Py_ssize_t)strlen(c->v.s_inline), "replace");
        Py_RETURN_NONE;
    default:
        Py_RETURN_NONE;
    }
}

/* ========== Cell find / set helpers ========== */

/* Key used for sorting and binary search: row-major */
static inline uint64_t cell_key(uint32_t row, uint16_t col) {
    return ((uint64_t)row << 16) | (uint64_t)col;
}

/* Binary search ws->cells for (row, col). Returns pointer or NULL. */
static OxlCell *ws_find_cell(OxlWorksheet *ws, uint32_t row, uint16_t col) {
    if (!ws->cells || ws->cell_count == 0) return NULL;
    uint64_t key = cell_key(row, col);
    uint32_t lo = 0, hi = ws->cell_count;
    while (lo < hi) {
        uint32_t mid = lo + (hi - lo) / 2;
        uint64_t mk = cell_key(ws->cells[mid].row, ws->cells[mid].col);
        if (mk == key) return &ws->cells[mid];
        if (mk < key)  lo = mid + 1;
        else           hi = mid;
    }
    return NULL;
}

/* Find insertion position for (row, col) — returns index where cell should go. */
static uint32_t ws_find_pos(OxlWorksheet *ws, uint32_t row, uint16_t col) {
    if (!ws->cells || ws->cell_count == 0) return 0;
    uint64_t key = cell_key(row, col);
    uint32_t lo = 0, hi = ws->cell_count;
    while (lo < hi) {
        uint32_t mid = lo + (hi - lo) / 2;
        uint64_t mk = cell_key(ws->cells[mid].row, ws->cells[mid].col);
        if (mk < key) lo = mid + 1;
        else          hi = mid;
    }
    return lo;
}

/* Free heap-owned fields from a cell before overwriting. */
static void cell_free_inline(OxlCell *c) {
    if (c->type == OXL_CELL_INLINE_STR || c->type == OXL_CELL_ERROR) {
        free(c->v.s_inline);
        c->v.s_inline = NULL;
    }
    free(c->formula);
    c->formula = NULL;
}

/* Set a cell at (row, col). If it exists, update type+value. If not, insert sorted.
   Returns 0 on success, -1 on error. */
static int ws_set_cell(OxlWorksheet *ws, OxlWorkbook *wb, uint32_t row, uint16_t col,
                       PyObject *value) {
    /* Build the cell value */
    OxlCell newc;
    memset(&newc, 0, sizeof(newc));
    newc.row = row;
    newc.col = col;

    if (value == Py_None) {
        newc.type = OXL_CELL_EMPTY;
    } else if (PyBool_Check(value)) {
        newc.type = OXL_CELL_BOOL;
        newc.v.b  = (value == Py_True) ? 1 : 0;
    } else if (PyLong_Check(value)) {
        newc.type = OXL_CELL_FLOAT;
        newc.v.f  = PyLong_AsDouble(value);
        if (newc.v.f == -1.0 && PyErr_Occurred()) return -1;
    } else if (PyFloat_Check(value)) {
        newc.type = OXL_CELL_FLOAT;
        newc.v.f  = PyFloat_AsDouble(value);
        if (newc.v.f == -1.0 && PyErr_Occurred()) return -1;
    } else if (PyUnicode_Check(value)) {
        const char *s = PyUnicode_AsUTF8(value);
        if (!s) return -1;
        if (s[0] == '=') {
            /* Formula: store raw expression (without '=') in cell->formula */
            char *dup = strdup(s + 1);
            if (!dup) { PyErr_NoMemory(); return -1; }
            newc.formula = dup;
            newc.type = OXL_CELL_EMPTY;  /* no cached numeric value */
        } else {
            uint32_t idx = oxl_sst_intern(&wb->sst, s);
            newc.type    = OXL_CELL_STRING;
            newc.v.s_idx = idx;
        }
    } else if (PyDateTime_Check(value)) {
        OxlDate d = {0};
        d.year  = (int16_t)PyDateTime_GET_YEAR(value);
        d.month = (uint8_t)PyDateTime_GET_MONTH(value);
        d.day   = (uint8_t)PyDateTime_GET_DAY(value);
        d.hour  = (uint8_t)PyDateTime_DATE_GET_HOUR(value);
        d.min   = (uint8_t)PyDateTime_DATE_GET_MINUTE(value);
        d.sec   = (uint8_t)PyDateTime_DATE_GET_SECOND(value);
        d.usec  = (uint32_t)PyDateTime_DATE_GET_MICROSECOND(value);
        newc.type = OXL_CELL_DATE;
        newc.v.dt = d;
    } else if (PyDate_Check(value)) {
        OxlDate d = {0};
        d.year  = (int16_t)PyDateTime_GET_YEAR(value);
        d.month = (uint8_t)PyDateTime_GET_MONTH(value);
        d.day   = (uint8_t)PyDateTime_GET_DAY(value);
        newc.type = OXL_CELL_DATE;
        newc.v.dt = d;
    } else {
        /* Fallback: convert to string */
        PyObject *str_obj = PyObject_Str(value);
        if (!str_obj) return -1;
        const char *s = PyUnicode_AsUTF8(str_obj);
        if (!s) { Py_DECREF(str_obj); return -1; }
        uint32_t idx = oxl_sst_intern(&wb->sst, s);
        Py_DECREF(str_obj);
        newc.type    = OXL_CELL_STRING;
        newc.v.s_idx = idx;
    }

    /* Try to find existing cell */
    uint32_t pos = ws_find_pos(ws, row, col);
    if (pos < ws->cell_count) {
        uint64_t ek = cell_key(ws->cells[pos].row, ws->cells[pos].col);
        if (ek == cell_key(row, col)) {
            /* Found: update in place, preserve style_idx */
            uint16_t saved_style = ws->cells[pos].style_idx;
            cell_free_inline(&ws->cells[pos]);
            ws->cells[pos].type    = newc.type;
            ws->cells[pos].v       = newc.v;
            ws->cells[pos].formula = newc.formula;
            if (ws->cells[pos].style_idx == 0)
                ws->cells[pos].style_idx = saved_style;
            return 0;
        }
    }

    /* Not found: insert at pos */
    if (newc.type == OXL_CELL_EMPTY && !newc.formula && newc.style_idx == 0) {
        /* Nothing to insert for truly empty cells that don't exist */
        return 0;
    }

    /* Grow if needed */
    if (ws->cell_count >= ws->cell_capacity) {
        uint32_t new_cap = ws->cell_capacity == 0 ? 8 : ws->cell_capacity * 2;
        OxlCell *new_cells = (OxlCell *)realloc(ws->cells, new_cap * sizeof(OxlCell));
        if (!new_cells) { PyErr_NoMemory(); return -1; }
        ws->cells = new_cells;
        ws->cell_capacity = new_cap;
    }

    /* Shift cells to make room */
    if (pos < ws->cell_count) {
        memmove(&ws->cells[pos + 1], &ws->cells[pos],
                (ws->cell_count - pos) * sizeof(OxlCell));
    }
    ws->cells[pos] = newc;
    ws->cell_count++;

    /* Update row_count and col_count */
    if (row + 1 > ws->row_count) ws->row_count = row + 1;
    if ((uint32_t)col + 1 > ws->col_count) ws->col_count = (uint32_t)col + 1;

    return 0;
}

/* ========== PyRowIteratorType ========== */

static PyObject *rowiter_next(PyRowIteratorObject *self) {
    OxlWorksheet *ws = self->ws;

    /* Skip rows before min_row */
    while (self->cell_pos < ws->cell_count &&
           ws->cells[self->cell_pos].row < self->min_row)
        self->cell_pos++;

    if (self->cell_pos >= ws->cell_count || self->cur_row > self->max_row)
        return NULL; /* StopIteration */

    /* Find the next row that has cells */
    uint32_t row = ws->cells[self->cell_pos].row;
    if (row > self->max_row) return NULL;

    /* If we skipped rows (empty rows), advance cur_row */
    self->cur_row = row;

    /* Count columns */
    uint16_t ncols = (uint16_t)(self->max_col - self->min_col + 1);
    PyObject *tup = PyTuple_New((Py_ssize_t)ncols);
    if (!tup) return NULL;

    /* Pre-fill with None */
    for (Py_ssize_t i = 0; i < (Py_ssize_t)ncols; i++) {
        Py_INCREF(Py_None);
        PyTuple_SET_ITEM(tup, i, Py_None);
    }

    /* Fill cells from this row */
    while (self->cell_pos < ws->cell_count &&
           ws->cells[self->cell_pos].row == row) {
        const OxlCell *c = &ws->cells[self->cell_pos];
        self->cell_pos++;
        if (c->col < self->min_col || c->col > self->max_col) continue;
        Py_ssize_t idx = (Py_ssize_t)(c->col - self->min_col);
        PyObject *val = cell_to_python(c, self->wb);
        if (!val) { Py_DECREF(tup); return NULL; }
        PyObject *old = PyTuple_GET_ITEM(tup, idx);
        Py_DECREF(old);
        PyTuple_SET_ITEM(tup, idx, val);
    }

    self->cur_row++;
    return tup;
}

static void rowiter_dealloc(PyRowIteratorObject *self) {
    Py_XDECREF(self->owner_ws);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyTypeObject PyRowIteratorType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.RowIterator",
    .tp_basicsize = sizeof(PyRowIteratorObject),
    .tp_dealloc   = (destructor)rowiter_dealloc,
    .tp_iternext  = (iternextfunc)rowiter_next,
    .tp_iter      = PyObject_SelfIter,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
};

/* ========== PyXlCellType ========== */

static void cell_dealloc(PyXlCellObject *self) {
    Py_XDECREF(self->owner_ws);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *cell_get_value(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    if (!c) Py_RETURN_NONE;
    return cell_to_python(c, self->wb);
}

static int cell_set_value(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value) {
        PyErr_SetString(PyExc_TypeError, "Cannot delete cell value");
        return -1;
    }
    return ws_set_cell(self->ws, self->wb, self->row, self->col, value);
}

static PyObject *cell_get_row(PyXlCellObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromUnsignedLong((unsigned long)(self->row + 1));
}

static PyObject *cell_get_column(PyXlCellObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromUnsignedLong((unsigned long)(self->col + 1));
}

static PyObject *cell_get_coordinate(PyXlCellObject *self, void *Py_UNUSED(x)) {
    char col_str[8];
    col_idx_to_str((int)(self->col + 1), col_str);
    char coord[32];
    snprintf(coord, sizeof(coord), "%s%u", col_str, self->row + 1);
    return PyUnicode_FromString(coord);
}

static PyObject *cell_get_column_letter(PyXlCellObject *self, void *Py_UNUSED(x)) {
    char col_str[8];
    col_idx_to_str((int)(self->col + 1), col_str);
    return PyUnicode_FromString(col_str);
}

static PyObject *cell_get_data_type(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    if (!c) return PyUnicode_FromString("n");
    if (c->formula) return PyUnicode_FromString("f");
    switch (c->type) {
    case OXL_CELL_EMPTY:      return PyUnicode_FromString("n");
    case OXL_CELL_FLOAT:      return PyUnicode_FromString("n");
    case OXL_CELL_STRING:     return PyUnicode_FromString("s");
    case OXL_CELL_INLINE_STR: return PyUnicode_FromString("s");
    case OXL_CELL_BOOL:       return PyUnicode_FromString("b");
    case OXL_CELL_DATE:       return PyUnicode_FromString("d");
    case OXL_CELL_ERROR:      return PyUnicode_FromString("e");
    default:                  return PyUnicode_FromString("n");
    }
}

static PyObject *cell_get_number_format(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    uint16_t style_idx = c ? c->style_idx : 0;
    const char *fmt = oxl_styles_get_numfmt_str(&self->wb->styles, style_idx);
    if (!fmt) Py_RETURN_NONE;
    return PyUnicode_FromString(fmt);
}

static int cell_set_number_format(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value || value == Py_None) {
        OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
        if (c) c->style_idx = 0;
        return 0;
    }
    if (!PyUnicode_Check(value)) {
        PyErr_SetString(PyExc_TypeError, "number_format must be a string");
        return -1;
    }
    const char *fmt_str = PyUnicode_AsUTF8(value);
    if (!fmt_str) return -1;
    oxl_styles_init_write_defaults(&self->wb->styles);
    uint16_t xf_idx = oxl_styles_get_or_add_xf(&self->wb->styles, fmt_str);

    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    if (c) {
        c->style_idx = xf_idx;
    } else {
        /* Create a stub EMPTY cell to carry the style */
        OxlCell newc;
        memset(&newc, 0, sizeof(newc));
        newc.row = self->row;
        newc.col = self->col;
        newc.style_idx = xf_idx;
        newc.type = OXL_CELL_EMPTY;
        uint32_t pos = ws_find_pos(self->ws, self->row, self->col);
        if (self->ws->cell_count >= self->ws->cell_capacity) {
            uint32_t new_cap = self->ws->cell_capacity == 0 ? 8 : self->ws->cell_capacity * 2;
            OxlCell *nc = (OxlCell *)realloc(self->ws->cells, new_cap * sizeof(OxlCell));
            if (!nc) { PyErr_NoMemory(); return -1; }
            self->ws->cells = nc;
            self->ws->cell_capacity = new_cap;
        }
        if (pos < self->ws->cell_count)
            memmove(&self->ws->cells[pos + 1], &self->ws->cells[pos],
                    (self->ws->cell_count - pos) * sizeof(OxlCell));
        self->ws->cells[pos] = newc;
        self->ws->cell_count++;
    }
    return 0;
}

static PyObject *cell_repr(PyXlCellObject *self) {
    char col_str[8];
    col_idx_to_str((int)(self->col + 1), col_str);
    char coord[32];
    snprintf(coord, sizeof(coord), "%s%u", col_str, self->row + 1);
    const char *sheet_name = self->ws->name ? self->ws->name : "";
    return PyUnicode_FromFormat("<Cell '%s'.%s>", sheet_name, coord);
}

static PyGetSetDef cell_getset[] = {
    {"value",         (getter)cell_get_value,          (setter)cell_set_value,          "Cell value", NULL},
    {"row",           (getter)cell_get_row,             NULL,                            "Row number (1-based)", NULL},
    {"column",        (getter)cell_get_column,          NULL,                            "Column number (1-based)", NULL},
    {"coordinate",    (getter)cell_get_coordinate,      NULL,                            "Cell coordinate (e.g. 'A1')", NULL},
    {"column_letter", (getter)cell_get_column_letter,   NULL,                            "Column letter (e.g. 'A')", NULL},
    {"data_type",     (getter)cell_get_data_type,       NULL,                            "Data type character", NULL},
    {"number_format", (getter)cell_get_number_format,   (setter)cell_set_number_format,  "Number format string", NULL},
    {NULL}
};

static PyTypeObject PyXlCellType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Cell",
    .tp_basicsize = sizeof(PyXlCellObject),
    .tp_dealloc   = (destructor)cell_dealloc,
    .tp_repr      = (reprfunc)cell_repr,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_getset    = cell_getset,
};

/* ========== PyWorksheetType ========== */

static PyRowIteratorObject *
make_row_iter(PyWorksheetObject *ws_obj, uint32_t min_row, uint32_t max_row,
              uint16_t min_col, uint16_t max_col) {
    PyRowIteratorObject *it = PyObject_New(PyRowIteratorObject, &PyRowIteratorType);
    if (!it) return NULL;
    Py_INCREF(ws_obj);
    it->owner_ws = (PyObject *)ws_obj;
    it->ws       = ws_obj->ws;
    it->wb       = ws_obj->wb;
    it->cell_pos = 0;
    it->cur_row  = min_row;
    it->min_row  = min_row;
    it->max_row  = max_row;
    it->min_col  = min_col;
    it->max_col  = max_col;
    return it;
}

static PyXlCellObject *make_cell_obj(PyWorksheetObject *ws_obj, uint32_t row, uint16_t col) {
    PyXlCellObject *obj = PyObject_New(PyXlCellObject, &PyXlCellType);
    if (!obj) return NULL;
    Py_INCREF(ws_obj);
    obj->owner_ws = (PyObject *)ws_obj;
    obj->ws = ws_obj->ws;
    obj->wb = ws_obj->wb;
    obj->row = row;
    obj->col = col;
    return obj;
}

static PyObject *worksheet_iter(PyWorksheetObject *self, PyObject *Py_UNUSED(args)) {
    OxlWorksheet *ws = self->ws;
    uint32_t max_row = ws->row_count > 0 ? ws->row_count - 1 : 0;
    uint16_t max_col = ws->col_count > 0 ? (uint16_t)(ws->col_count - 1) : 0;
    return (PyObject *)make_row_iter(self, 0, max_row, 0, max_col);
}

static PyObject *worksheet_iter_rows(PyWorksheetObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"min_row","max_row","min_col","max_col", NULL};
    int min_row = 0, max_row = -1, min_col = 0, max_col = -1;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|iiii", kwlist,
                                     &min_row, &max_row, &min_col, &max_col))
        return NULL;
    OxlWorksheet *ws = self->ws;
    if (max_row < 0) max_row = ws->row_count > 0 ? (int)ws->row_count - 1 : 0;
    if (max_col < 0) max_col = ws->col_count > 0 ? (int)ws->col_count - 1 : 0;
    return (PyObject *)make_row_iter(self,
        (uint32_t)min_row, (uint32_t)max_row,
        (uint16_t)min_col, (uint16_t)max_col);
}

static PyObject *worksheet_append(PyWorksheetObject *self, PyObject *arg) {
    PyObject *seq = PySequence_Fast(arg, "append() requires a sequence");
    if (!seq) return NULL;
    Py_ssize_t n = PySequence_Fast_GET_SIZE(seq);
    uint32_t row = self->ws->row_count;

    for (Py_ssize_t i = 0; i < n; i++) {
        PyObject *item = PySequence_Fast_GET_ITEM(seq, i);
        OxlCell c;
        memset(&c, 0, sizeof(c));
        c.row = row;
        c.col = (uint16_t)i;

        if (item == Py_None) {
            c.type = OXL_CELL_EMPTY;
        } else if (PyBool_Check(item)) {
            c.type = OXL_CELL_BOOL;
            c.v.b  = (item == Py_True) ? 1 : 0;
        } else if (PyLong_Check(item)) {
            c.type = OXL_CELL_FLOAT;
            c.v.f  = PyLong_AsDouble(item);
        } else if (PyFloat_Check(item)) {
            c.type = OXL_CELL_FLOAT;
            c.v.f  = PyFloat_AsDouble(item);
        } else if (PyUnicode_Check(item)) {
            const char *s = PyUnicode_AsUTF8(item);
            if (!s) { Py_DECREF(seq); return NULL; }
            uint32_t idx = oxl_sst_intern(&self->wb->sst, s);
            c.type    = OXL_CELL_STRING;
            c.v.s_idx = idx;
        } else if (PyDate_Check(item) && !PyDateTime_Check(item)) {
            OxlDate d = {0};
            d.year  = (int16_t)PyDateTime_GET_YEAR(item);
            d.month = (uint8_t)PyDateTime_GET_MONTH(item);
            d.day   = (uint8_t)PyDateTime_GET_DAY(item);
            c.type = OXL_CELL_DATE;
            c.v.dt = d;
        } else if (PyDateTime_Check(item)) {
            OxlDate d = {0};
            d.year  = (int16_t)PyDateTime_GET_YEAR(item);
            d.month = (uint8_t)PyDateTime_GET_MONTH(item);
            d.day   = (uint8_t)PyDateTime_GET_DAY(item);
            d.hour  = (uint8_t)PyDateTime_DATE_GET_HOUR(item);
            d.min   = (uint8_t)PyDateTime_DATE_GET_MINUTE(item);
            d.sec   = (uint8_t)PyDateTime_DATE_GET_SECOND(item);
            c.type = OXL_CELL_DATE;
            c.v.dt = d;
        } else {
            /* Fallback: str(item) */
            PyObject *s = PyObject_Str(item);
            if (!s) { Py_DECREF(seq); return NULL; }
            const char *cs = PyUnicode_AsUTF8(s);
            if (cs) {
                uint32_t idx = oxl_sst_intern(&self->wb->sst, cs);
                c.type    = OXL_CELL_STRING;
                c.v.s_idx = idx;
            }
            Py_DECREF(s);
        }

        if (c.type != OXL_CELL_EMPTY)
            oxl_worksheet_append_cell(self->ws, &c);
    }
    Py_DECREF(seq);
    Py_RETURN_NONE;
}

/* worksheet_subscript: supports ws['A1'], ws['A1:C3'], ws[(row, col)] */
static PyObject *worksheet_subscript(PyWorksheetObject *self, PyObject *key) {
    if (PyUnicode_Check(key)) {
        const char *ref = PyUnicode_AsUTF8(key);
        if (!ref) return NULL;

        /* Check for range: contains ':' */
        const char *colon = strchr(ref, ':');
        if (colon) {
            /* Range: parse "A1:C3" */
            /* Split at colon */
            size_t start_len = (size_t)(colon - ref);
            char start_ref[32], end_ref[32];
            if (start_len >= sizeof(start_ref) || strlen(colon + 1) >= sizeof(end_ref)) {
                PyErr_SetString(PyExc_ValueError, "Cell reference too long");
                return NULL;
            }
            memcpy(start_ref, ref, start_len);
            start_ref[start_len] = '\0';
            strcpy(end_ref, colon + 1);

            uint32_t min_row, max_row;
            uint16_t min_col, max_col;
            if (parse_a1_ref(start_ref, &min_row, &min_col) < 0 ||
                parse_a1_ref(end_ref, &max_row, &max_col) < 0) {
                PyErr_Format(PyExc_ValueError, "Invalid range reference: %s", ref);
                return NULL;
            }

            /* Ensure min <= max */
            if (min_row > max_row) { uint32_t t = min_row; min_row = max_row; max_row = t; }
            if (min_col > max_col) { uint16_t t = min_col; min_col = max_col; max_col = t; }

            uint32_t nrows = max_row - min_row + 1;
            uint32_t ncols = (uint32_t)(max_col - min_col + 1);

            PyObject *outer = PyTuple_New((Py_ssize_t)nrows);
            if (!outer) return NULL;

            for (uint32_t r = 0; r < nrows; r++) {
                PyObject *inner = PyTuple_New((Py_ssize_t)ncols);
                if (!inner) { Py_DECREF(outer); return NULL; }
                for (uint32_t c = 0; c < ncols; c++) {
                    PyXlCellObject *cell = make_cell_obj(self,
                        min_row + r, (uint16_t)(min_col + c));
                    if (!cell) { Py_DECREF(inner); Py_DECREF(outer); return NULL; }
                    PyTuple_SET_ITEM(inner, (Py_ssize_t)c, (PyObject *)cell);
                }
                PyTuple_SET_ITEM(outer, (Py_ssize_t)r, inner);
            }
            return outer;
        } else {
            /* Single cell reference */
            uint32_t row;
            uint16_t col;
            if (parse_a1_ref(ref, &row, &col) < 0) {
                PyErr_Format(PyExc_ValueError, "Invalid cell reference: %s", ref);
                return NULL;
            }
            return (PyObject *)make_cell_obj(self, row, col);
        }
    } else if (PyTuple_Check(key) && PyTuple_Size(key) == 2) {
        /* Tuple (row, col), 1-based */
        PyObject *row_obj = PyTuple_GET_ITEM(key, 0);
        PyObject *col_obj = PyTuple_GET_ITEM(key, 1);
        if (!PyLong_Check(row_obj) || !PyLong_Check(col_obj)) {
            PyErr_SetString(PyExc_TypeError, "Tuple key must be (int, int)");
            return NULL;
        }
        long r = PyLong_AsLong(row_obj);
        long c = PyLong_AsLong(col_obj);
        if (r < 1 || c < 1) {
            PyErr_SetString(PyExc_ValueError, "Row and column must be >= 1");
            return NULL;
        }
        return (PyObject *)make_cell_obj(self, (uint32_t)(r - 1), (uint16_t)(c - 1));
    } else {
        PyErr_SetString(PyExc_TypeError, "Worksheet key must be a string or (row, col) tuple");
        return NULL;
    }
}

static PyObject *worksheet_cell(PyWorksheetObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"row", "column", NULL};
    int row = 1, col = 1;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|ii", kwlist, &row, &col))
        return NULL;
    if (row < 1 || col < 1) {
        PyErr_SetString(PyExc_ValueError, "row and column must be >= 1");
        return NULL;
    }
    return (PyObject *)make_cell_obj(self, (uint32_t)(row - 1), (uint16_t)(col - 1));
}

static PyObject *worksheet_get_max_row(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromUnsignedLong(self->ws->row_count);
}

static PyObject *worksheet_get_max_col(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromUnsignedLong(self->ws->col_count);
}

static PyObject *worksheet_get_title(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    if (!self->ws->name) Py_RETURN_NONE;
    return PyUnicode_DecodeUTF8(self->ws->name, (Py_ssize_t)strlen(self->ws->name), "replace");
}

/* ── Phase 5: Merged cells ──────────────────────────────────────────────── */

static int parse_range_ref(const char *ref, uint32_t *min_row, uint16_t *min_col,
                            uint32_t *max_row, uint16_t *max_col) {
    const char *colon = strchr(ref, ':');
    if (!colon) return -1;
    char start[32], end[32];
    size_t slen = (size_t)(colon - ref);
    if (slen >= sizeof(start) || strlen(colon + 1) >= sizeof(end)) return -1;
    memcpy(start, ref, slen); start[slen] = '\0';
    strcpy(end, colon + 1);
    if (parse_a1_ref(start, min_row, min_col) < 0) return -1;
    if (parse_a1_ref(end,   max_row, max_col) < 0) return -1;
    return 0;
}

static PyObject *worksheet_merge_cells(PyWorksheetObject *self, PyObject *args) {
    const char *ref;
    if (!PyArg_ParseTuple(args, "s", &ref)) return NULL;
    uint32_t r1, r2; uint16_t c1, c2;
    if (parse_range_ref(ref, &r1, &c1, &r2, &c2) < 0) {
        PyErr_Format(PyExc_ValueError, "Invalid range: %s", ref);
        return NULL;
    }
    if (oxl_worksheet_add_merge(self->ws, r1 + 1, c1 + 1, r2 + 1, c2 + 1) < 0)
        return PyErr_NoMemory();
    Py_RETURN_NONE;
}

static PyObject *worksheet_unmerge_cells(PyWorksheetObject *self, PyObject *args) {
    const char *ref;
    if (!PyArg_ParseTuple(args, "s", &ref)) return NULL;
    uint32_t r1, r2; uint16_t c1, c2;
    if (parse_range_ref(ref, &r1, &c1, &r2, &c2) < 0) {
        PyErr_Format(PyExc_ValueError, "Invalid range: %s", ref);
        return NULL;
    }
    /* Convert to 1-based to match storage */
    uint32_t mr1 = r1 + 1, mr2 = r2 + 1;
    uint16_t mc1 = c1 + 1, mc2 = c2 + 1;
    OxlWorksheet *ws = self->ws;
    for (uint32_t i = 0; i < ws->merged_cell_count; i++) {
        OxlMergedCell *m = &ws->merged_cells[i];
        if (m->min_row == mr1 && m->min_col == mc1 &&
            m->max_row == mr2 && m->max_col == mc2) {
            memmove(&ws->merged_cells[i], &ws->merged_cells[i + 1],
                    (ws->merged_cell_count - i - 1) * sizeof(OxlMergedCell));
            ws->merged_cell_count--;
            Py_RETURN_NONE;
        }
    }
    Py_RETURN_NONE;  /* silently ignore if not found */
}

static PyObject *worksheet_get_merged_cells(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    OxlWorksheet *ws = self->ws;
    PyObject *tup = PyTuple_New((Py_ssize_t)ws->merged_cell_count);
    if (!tup) return NULL;
    char buf[64];
    for (uint32_t i = 0; i < ws->merged_cell_count; i++) {
        const OxlMergedCell *m = &ws->merged_cells[i];
        char col1[5], col2[5];
        col_idx_to_str((int)m->min_col, col1);
        col_idx_to_str((int)m->max_col, col2);
        snprintf(buf, sizeof(buf), "%s%u:%s%u", col1, m->min_row, col2, m->max_row);
        PyObject *s = PyUnicode_FromString(buf);
        if (!s) { Py_DECREF(tup); return NULL; }
        PyTuple_SET_ITEM(tup, (Py_ssize_t)i, s);
    }
    return tup;
}

/* ── Phase 4: Column/Row Dimensions ─────────────────────────────────────── */

static PyObject *worksheet_set_column_width(PyWorksheetObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"col", "width", "hidden", NULL};
    PyObject *col_obj;
    double width = 0.0;
    int hidden = 0;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "Od|i", kwlist, &col_obj, &width, &hidden))
        return NULL;
    int ci;
    if (PyLong_Check(col_obj)) {
        ci = (int)PyLong_AsLong(col_obj);
    } else if (PyUnicode_Check(col_obj)) {
        const char *s = PyUnicode_AsUTF8(col_obj);
        if (!s) return NULL;
        ci = col_str_to_idx(s);
    } else {
        PyErr_SetString(PyExc_TypeError, "col must be str or int");
        return NULL;
    }
    if (ci < 1) { PyErr_SetString(PyExc_ValueError, "Invalid column"); return NULL; }
    int custom = width > 0.0;
    if (oxl_worksheet_set_col_dim(self->ws, (uint16_t)ci, (uint16_t)ci,
                                   width, hidden, 0, custom) < 0)
        return PyErr_NoMemory();
    Py_RETURN_NONE;
}

static PyObject *worksheet_set_row_height(PyWorksheetObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"row", "height", "hidden", NULL};
    int row;
    double height = 0.0;
    int hidden = 0;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "id|i", kwlist, &row, &height, &hidden))
        return NULL;
    if (row < 1) { PyErr_SetString(PyExc_ValueError, "Row must be >= 1"); return NULL; }
    int custom = height > 0.0;
    if (oxl_worksheet_set_row_dim(self->ws, (uint32_t)row, height, hidden, custom) < 0)
        return PyErr_NoMemory();
    Py_RETURN_NONE;
}

/* ── Phase 9: Freeze panes / sheet view ──────────────────────────────────── */

static PyObject *worksheet_get_freeze_panes(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    if (!self->ws->freeze_panes) Py_RETURN_NONE;
    return PyUnicode_FromString(self->ws->freeze_panes);
}

static int worksheet_set_freeze_panes(PyWorksheetObject *self, PyObject *v, void *Py_UNUSED(x)) {
    free(self->ws->freeze_panes);
    self->ws->freeze_panes = NULL;
    if (v == Py_None) return 0;
    const char *s = PyUnicode_AsUTF8(v);
    if (!s) return -1;
    self->ws->freeze_panes = strdup(s);
    return self->ws->freeze_panes ? 0 : (PyErr_NoMemory(), -1);
}

static PyObject *worksheet_get_zoom_scale(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromLong(self->ws->zoom_scale);
}

static int worksheet_set_zoom_scale(PyWorksheetObject *self, PyObject *v, void *Py_UNUSED(x)) {
    long z = PyLong_AsLong(v);
    if (z == -1 && PyErr_Occurred()) return -1;
    self->ws->zoom_scale = (int)z;
    return 0;
}

static PyObject *worksheet_get_show_gridlines(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->ws->show_gridlines);
}

static int worksheet_set_show_gridlines(PyWorksheetObject *self, PyObject *v, void *Py_UNUSED(x)) {
    self->ws->show_gridlines = PyObject_IsTrue(v);
    return 0;
}

/* ── Phase 10: Tab color ─────────────────────────────────────────────────── */

static PyObject *worksheet_get_tab_color(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    if (!self->ws->tab_color[0]) Py_RETURN_NONE;
    return PyUnicode_FromString(self->ws->tab_color);
}

static int worksheet_set_tab_color(PyWorksheetObject *self, PyObject *v, void *Py_UNUSED(x)) {
    if (v == Py_None) { self->ws->tab_color[0] = '\0'; return 0; }
    const char *s = PyUnicode_AsUTF8(v);
    if (!s) return -1;
    strncpy(self->ws->tab_color, s, 8);
    self->ws->tab_color[8] = '\0';
    return 0;
}

/* ── Phase 11: Auto-filter ───────────────────────────────────────────────── */

static PyObject *worksheet_get_auto_filter_ref(PyWorksheetObject *self, void *Py_UNUSED(x)) {
    if (!self->ws->auto_filter_ref) Py_RETURN_NONE;
    return PyUnicode_FromString(self->ws->auto_filter_ref);
}

static int worksheet_set_auto_filter_ref(PyWorksheetObject *self, PyObject *v, void *Py_UNUSED(x)) {
    free(self->ws->auto_filter_ref);
    self->ws->auto_filter_ref = NULL;
    if (v == Py_None) return 0;
    const char *s = PyUnicode_AsUTF8(v);
    if (!s) return -1;
    self->ws->auto_filter_ref = strdup(s);
    return self->ws->auto_filter_ref ? 0 : (PyErr_NoMemory(), -1);
}

static void worksheet_dealloc(PyWorksheetObject *self) {
    Py_XDECREF(self->owner);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyMethodDef worksheet_methods[] = {
    {"__iter__",         (PyCFunction)worksheet_iter,           METH_NOARGS,   "Iterate rows"},
    {"iter_rows",        (PyCFunction)(void(*)(void))worksheet_iter_rows, METH_VARARGS|METH_KEYWORDS, "iter_rows(min_row=0, max_row=-1, min_col=0, max_col=-1)"},
    {"append",           (PyCFunction)worksheet_append,         METH_O,        "Append a row"},
    {"cell",             (PyCFunction)(void(*)(void))worksheet_cell, METH_VARARGS|METH_KEYWORDS, "cell(row=1, column=1)"},
    {"merge_cells",      (PyCFunction)worksheet_merge_cells,    METH_VARARGS,  "merge_cells('A1:C3')"},
    {"unmerge_cells",    (PyCFunction)worksheet_unmerge_cells,  METH_VARARGS,  "unmerge_cells('A1:C3')"},
    {"set_column_width", (PyCFunction)(void(*)(void))worksheet_set_column_width, METH_VARARGS|METH_KEYWORDS, "set_column_width(col, width, hidden=False)"},
    {"set_row_height",   (PyCFunction)(void(*)(void))worksheet_set_row_height,   METH_VARARGS|METH_KEYWORDS, "set_row_height(row, height, hidden=False)"},
    {NULL, NULL}
};

static PyGetSetDef worksheet_getset[] = {
    {"max_row",         (getter)worksheet_get_max_row,         NULL,                           "Number of rows",     NULL},
    {"max_column",      (getter)worksheet_get_max_col,         NULL,                           "Number of columns",  NULL},
    {"title",           (getter)worksheet_get_title,           NULL,                           "Sheet name",         NULL},
    {"merged_cells",    (getter)worksheet_get_merged_cells,    NULL,                           "Merged cell ranges", NULL},
    {"freeze_panes",    (getter)worksheet_get_freeze_panes,    (setter)worksheet_set_freeze_panes,    "Freeze panes cell ref", NULL},
    {"zoom_scale",      (getter)worksheet_get_zoom_scale,      (setter)worksheet_set_zoom_scale,      "Zoom scale %",       NULL},
    {"show_gridlines",  (getter)worksheet_get_show_gridlines,  (setter)worksheet_set_show_gridlines,  "Show gridlines",     NULL},
    {"tab_color",       (getter)worksheet_get_tab_color,       (setter)worksheet_set_tab_color,       "Tab color hex",      NULL},
    {"auto_filter_ref", (getter)worksheet_get_auto_filter_ref, (setter)worksheet_set_auto_filter_ref, "Auto-filter range",  NULL},
    {NULL}
};

static PyMappingMethods worksheet_mapping = {
    .mp_subscript = (binaryfunc)worksheet_subscript,
};

static PyTypeObject PyWorksheetType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Worksheet",
    .tp_basicsize = sizeof(PyWorksheetObject),
    .tp_dealloc   = (destructor)worksheet_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_methods   = worksheet_methods,
    .tp_getset    = worksheet_getset,
    .tp_iter      = (getiterfunc)(void(*)(void))worksheet_iter,
    .tp_as_mapping = &worksheet_mapping,
};

/* ========== PyWorkbookType ========== */

static PyWorksheetObject *
make_worksheet_obj(PyWorkbookObject *wb_obj, OxlWorksheet *ws) {
    PyWorksheetObject *obj = PyObject_New(PyWorksheetObject, &PyWorksheetType);
    if (!obj) return NULL;
    Py_INCREF(wb_obj);
    obj->owner = (PyObject *)wb_obj;
    obj->ws    = ws;
    obj->wb    = wb_obj->wb;
    return obj;
}

static PyObject *workbook_create_sheet(PyWorkbookObject *self, PyObject *args) {
    const char *name = "Sheet";
    if (!PyArg_ParseTuple(args, "|s", &name)) return NULL;
    OxlWorksheet *ws = oxl_worksheet_new(name, NULL);
    if (!ws) return PyErr_NoMemory();
    if (oxl_workbook_add_sheet(self->wb, ws) < 0) {
        oxl_worksheet_free(ws);
        PyErr_SetString(PyExc_MemoryError, "Cannot add sheet");
        return NULL;
    }
    return (PyObject *)make_worksheet_obj(self, ws);
}

static PyObject *workbook_getitem(PyWorkbookObject *self, PyObject *key) {
    if (PyLong_Check(key)) {
        long idx = PyLong_AsLong(key);
        OxlWorksheet *ws = oxl_workbook_get_sheet_by_index(self->wb, (uint32_t)idx);
        if (!ws) { PyErr_SetString(PyExc_IndexError, "Sheet index out of range"); return NULL; }
        return (PyObject *)make_worksheet_obj(self, ws);
    }
    const char *name = PyUnicode_AsUTF8(key);
    if (!name) return NULL;
    OxlWorksheet *ws = oxl_workbook_get_sheet_by_name(self->wb, name);
    if (!ws) { PyErr_SetString(PyExc_KeyError, name); return NULL; }
    return (PyObject *)make_worksheet_obj(self, ws);
}

static PyObject *workbook_get_active(PyWorkbookObject *self, void *Py_UNUSED(x)) {
    if (self->wb->sheet_count == 0) Py_RETURN_NONE;
    return (PyObject *)make_worksheet_obj(self, self->wb->sheets[0]);
}

static PyObject *workbook_len(PyWorkbookObject *self, PyObject *Py_UNUSED(args)) {
    return PyLong_FromUnsignedLong(self->wb->sheet_count);
}

static PyObject *workbook_save(PyWorkbookObject *self, PyObject *args) {
    const char *path;
    if (!PyArg_ParseTuple(args, "s", &path)) return NULL;
    char err[256] = {0};
    int rc;
    Py_BEGIN_ALLOW_THREADS
    rc = oxl_write_workbook(self->wb, path, err, sizeof(err));
    Py_END_ALLOW_THREADS
    if (rc != 0) {
        PyErr_SetString(PyExc_IOError, err[0] ? err : "Failed to write workbook");
        return NULL;
    }
    Py_RETURN_NONE;
}

static PyObject *workbook_close(PyWorkbookObject *self, PyObject *Py_UNUSED(args)) {
    Py_RETURN_NONE;
}

static PyObject *workbook_enter(PyWorkbookObject *self, PyObject *Py_UNUSED(args)) {
    Py_INCREF(self);
    return (PyObject *)self;
}

static PyObject *workbook_exit(PyWorkbookObject *self, PyObject *args) {
    (void)args;
    return workbook_close(self, NULL);
}

static void workbook_dealloc(PyWorkbookObject *self) {
    oxl_workbook_free(self->wb);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *workbook_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    OxlWorkbook *wb = oxl_workbook_new();
    if (!wb) return PyErr_NoMemory();
    PyWorkbookObject *obj = (PyWorkbookObject *)type->tp_alloc(type, 0);
    if (!obj) { oxl_workbook_free(wb); return NULL; }
    obj->wb = wb;
    return (PyObject *)obj;
}

/* ── Phase 6: Defined names ──────────────────────────────────────────────── */

static PyObject *workbook_add_defined_name(PyWorkbookObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"name", "value", "local_sheet_id", "hidden", NULL};
    const char *name, *value;
    int local_sheet_id = -1, hidden = 0;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "ss|ii", kwlist,
                                     &name, &value, &local_sheet_id, &hidden))
        return NULL;
    if (oxl_workbook_add_defined_name(self->wb, name, value, local_sheet_id, hidden) < 0)
        return PyErr_NoMemory();
    Py_RETURN_NONE;
}

static PyObject *workbook_get_defined_names(PyWorkbookObject *self, void *Py_UNUSED(x)) {
    PyObject *d = PyDict_New();
    if (!d) return NULL;
    OxlWorkbook *wb = self->wb;
    for (uint32_t i = 0; i < wb->defined_name_count; i++) {
        const OxlDefinedName *dn = &wb->defined_names[i];
        if (!dn->name) continue;
        PyObject *key = PyUnicode_FromString(dn->name);
        PyObject *val = PyUnicode_FromString(dn->value ? dn->value : "");
        if (!key || !val) { Py_XDECREF(key); Py_XDECREF(val); Py_DECREF(d); return NULL; }
        PyDict_SetItem(d, key, val);
        Py_DECREF(key); Py_DECREF(val);
    }
    return d;
}

static PyMethodDef workbook_methods[] = {
    {"create_sheet",      (PyCFunction)workbook_create_sheet,  METH_VARARGS, "create_sheet(name='Sheet')"},
    {"save",              (PyCFunction)workbook_save,          METH_VARARGS, "save(path)"},
    {"close",             (PyCFunction)workbook_close,         METH_NOARGS,  "close()"},
    {"__enter__",         (PyCFunction)workbook_enter,         METH_NOARGS,  NULL},
    {"__exit__",          (PyCFunction)workbook_exit,          METH_VARARGS, NULL},
    {"__len__",           (PyCFunction)workbook_len,           METH_NOARGS,  NULL},
    {"add_defined_name",  (PyCFunction)(void(*)(void))workbook_add_defined_name, METH_VARARGS|METH_KEYWORDS, "add_defined_name(name, value, local_sheet_id=-1, hidden=0)"},
    {NULL, NULL}
};

static PyGetSetDef workbook_getset[] = {
    {"active",        (getter)workbook_get_active,        NULL, "First sheet",    NULL},
    {"defined_names", (getter)workbook_get_defined_names, NULL, "Defined names dict", NULL},
    {NULL}
};

static PyMappingMethods workbook_mapping = {
    .mp_subscript = (binaryfunc)workbook_getitem,
};

static PyTypeObject PyWorkbookType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Workbook",
    .tp_basicsize = sizeof(PyWorkbookObject),
    .tp_dealloc   = (destructor)workbook_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT | Py_TPFLAGS_BASETYPE,
    .tp_new       = workbook_new,
    .tp_methods   = workbook_methods,
    .tp_getset    = workbook_getset,
    .tp_as_mapping = &workbook_mapping,
};

/* ========== Module-level load_workbook ========== */

static PyObject *m_load_workbook(PyObject *Py_UNUSED(mod), PyObject *args) {
    const char *path;
    if (!PyArg_ParseTuple(args, "s", &path)) return NULL;

    char err[512] = {0};
    OxlWorkbook *wb = NULL;
    Py_BEGIN_ALLOW_THREADS
    wb = oxl_read_workbook(path, err, sizeof(err));
    Py_END_ALLOW_THREADS

    if (!wb) {
        PyErr_SetString(PyExc_IOError, err[0] ? err : "Failed to open workbook");
        return NULL;
    }
    PyWorkbookObject *obj = (PyWorkbookObject *)PyWorkbookType.tp_alloc(&PyWorkbookType, 0);
    if (!obj) { oxl_workbook_free(wb); return NULL; }
    obj->wb = wb;
    return (PyObject *)obj;
}

/* ========== Module-level column utilities ========== */

static PyObject *m_column_index_from_string(PyObject *Py_UNUSED(mod), PyObject *args) {
    const char *s;
    if (!PyArg_ParseTuple(args, "s", &s)) return NULL;
    /* Convert to uppercase */
    char upper[16];
    size_t i;
    for (i = 0; s[i] && i < sizeof(upper) - 1; i++)
        upper[i] = (char)(s[i] >= 'a' && s[i] <= 'z' ? s[i] - 32 : s[i]);
    upper[i] = '\0';
    int idx = col_str_to_idx(upper);
    if (idx <= 0) {
        PyErr_Format(PyExc_ValueError, "Invalid column string: %s", s);
        return NULL;
    }
    return PyLong_FromLong(idx);
}

static PyObject *m_get_column_letter(PyObject *Py_UNUSED(mod), PyObject *args) {
    int idx;
    if (!PyArg_ParseTuple(args, "i", &idx)) return NULL;
    if (idx <= 0) {
        PyErr_Format(PyExc_ValueError, "Column index must be >= 1, got %d", idx);
        return NULL;
    }
    char buf[8];
    col_idx_to_str(idx, buf);
    return PyUnicode_FromString(buf);
}

/* ========== Module definition ========== */

static PyMethodDef module_methods[] = {
    {"load_workbook",             m_load_workbook,             METH_VARARGS, "load_workbook(path) -> Workbook"},
    {"column_index_from_string",  m_column_index_from_string,  METH_VARARGS, "column_index_from_string(s) -> int"},
    {"get_column_letter",         m_get_column_letter,         METH_VARARGS, "get_column_letter(idx) -> str"},
    {NULL, NULL}
};

static struct PyModuleDef moduledef = {
    PyModuleDef_HEAD_INIT,
    "_openexcel",
    "High-performance .xlsx reader/writer",
    -1,
    module_methods,
};

PyMODINIT_FUNC PyInit__openexcel(void) {
    PyDateTime_IMPORT;

    if (PyType_Ready(&PyRowIteratorType) < 0) return NULL;
    if (PyType_Ready(&PyXlCellType) < 0)        return NULL;
    if (PyType_Ready(&PyWorksheetType) < 0)   return NULL;
    if (PyType_Ready(&PyWorkbookType) < 0)    return NULL;

    PyObject *mod = PyModule_Create(&moduledef);
    if (!mod) return NULL;

    Py_INCREF(&PyWorkbookType);
    if (PyModule_AddObject(mod, "Workbook", (PyObject *)&PyWorkbookType) < 0) {
        Py_DECREF(&PyWorkbookType);
        Py_DECREF(mod);
        return NULL;
    }

    Py_INCREF(&PyXlCellType);
    if (PyModule_AddObject(mod, "Cell", (PyObject *)&PyXlCellType) < 0) {
        Py_DECREF(&PyXlCellType);
        Py_DECREF(mod);
        return NULL;
    }

    return mod;
}
