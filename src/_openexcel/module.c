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
static PyTypeObject PyFontType;
static PyTypeObject PyPatternFillType;
static PyTypeObject PySideType;
static PyTypeObject PyBorderType;
static PyTypeObject PyAlignmentType;
static PyTypeObject PyXlDataValidationType;

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

/* ========== Phase 3: Style Python types ========== */

/* ---- Color helpers ---- */

static uint32_t parse_color_str(const char *s) {
    if (!s || !*s) return 0;
    if (*s == '#') s++;
    size_t len = strlen(s);
    unsigned long val = strtoul(s, NULL, 16);
    if (len <= 6) {
        val |= 0xFF000000UL;
    }
    return (uint32_t)val;
}

static PyObject *color_to_pystr(uint32_t argb) {
    if (argb == 0) Py_RETURN_NONE;
    char buf[9];
    snprintf(buf, sizeof(buf), "%08X", argb);
    return PyUnicode_FromString(buf);
}

/* ---- PyFontObject ---- */

typedef struct {
    PyObject_HEAD
    char    *name;
    double   size;
    char    *color;
    uint8_t  bold;
    uint8_t  italic;
    uint8_t  underline; /* 0=none, 1=single, 2=double */
} PyFontObject;

static void font_dealloc(PyFontObject *self) {
    free(self->name);
    free(self->color);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *font_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    PyFontObject *self = (PyFontObject *)type->tp_alloc(type, 0);
    if (self) {
        self->name = NULL;
        self->size = 11.0;
        self->color = NULL;
        self->bold = 0;
        self->italic = 0;
        self->underline = 0;
    }
    return (PyObject *)self;
}

static int font_init(PyFontObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"name","size","bold","italic","underline","color", NULL};
    const char *name = NULL;
    double size = 11.0;
    int bold = 0, italic = 0;
    PyObject *underline_obj = Py_None;
    PyObject *color_obj = Py_None;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|zdiiOO", kwlist,
                                     &name, &size, &bold, &italic,
                                     &underline_obj, &color_obj))
        return -1;

    free(self->name);
    self->name = name ? strdup(name) : NULL;
    self->size = size;
    self->bold = (uint8_t)(bold ? 1 : 0);
    self->italic = (uint8_t)(italic ? 1 : 0);

    /* underline: None -> 0, "single" -> 1, "double" -> 2 */
    if (underline_obj == Py_None) {
        self->underline = 0;
    } else if (PyUnicode_Check(underline_obj)) {
        const char *u = PyUnicode_AsUTF8(underline_obj);
        if (!u) return -1;
        if (strcmp(u, "single") == 0) self->underline = 1;
        else if (strcmp(u, "double") == 0) self->underline = 2;
        else self->underline = 1; /* default to single */
    } else {
        PyErr_SetString(PyExc_TypeError, "underline must be None, 'single', or 'double'");
        return -1;
    }

    free(self->color);
    self->color = NULL;
    if (color_obj != Py_None) {
        if (!PyUnicode_Check(color_obj)) {
            PyErr_SetString(PyExc_TypeError, "color must be a string or None");
            return -1;
        }
        const char *c = PyUnicode_AsUTF8(color_obj);
        if (!c) return -1;
        self->color = strdup(c);
        if (!self->color) { PyErr_NoMemory(); return -1; }
    }
    return 0;
}

static PyObject *font_get_name(PyFontObject *self, void *Py_UNUSED(x)) {
    if (!self->name) Py_RETURN_NONE;
    return PyUnicode_FromString(self->name);
}
static PyObject *font_get_size(PyFontObject *self, void *Py_UNUSED(x)) {
    return PyFloat_FromDouble(self->size);
}
static PyObject *font_get_bold(PyFontObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->bold);
}
static PyObject *font_get_italic(PyFontObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->italic);
}
static PyObject *font_get_underline(PyFontObject *self, void *Py_UNUSED(x)) {
    if (self->underline == 0) Py_RETURN_NONE;
    if (self->underline == 2) return PyUnicode_FromString("double");
    return PyUnicode_FromString("single");
}
static PyObject *font_get_color(PyFontObject *self, void *Py_UNUSED(x)) {
    if (!self->color) Py_RETURN_NONE;
    return PyUnicode_FromString(self->color);
}

static PyObject *font_repr(PyFontObject *self) {
    char buf[256];
    snprintf(buf, sizeof(buf), "<Font name='%s' size=%.6g bold=%s>",
        self->name ? self->name : "",
        self->size,
        self->bold ? "True" : "False");
    return PyUnicode_FromString(buf);
}

static PyGetSetDef font_getset[] = {
    {"name",      (getter)font_get_name,      NULL, "Font name", NULL},
    {"size",      (getter)font_get_size,      NULL, "Font size", NULL},
    {"bold",      (getter)font_get_bold,      NULL, "Bold", NULL},
    {"italic",    (getter)font_get_italic,    NULL, "Italic", NULL},
    {"underline", (getter)font_get_underline, NULL, "Underline", NULL},
    {"color",     (getter)font_get_color,     NULL, "Color ARGB hex", NULL},
    {NULL}
};

static PyTypeObject PyFontType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Font",
    .tp_basicsize = sizeof(PyFontObject),
    .tp_dealloc   = (destructor)font_dealloc,
    .tp_repr      = (reprfunc)font_repr,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_new       = font_new,
    .tp_init      = (initproc)font_init,
    .tp_getset    = font_getset,
};

/* ---- PyPatternFillObject ---- */

typedef struct {
    PyObject_HEAD
    char *fill_type;
    char *fg_color;
    char *bg_color;
} PyPatternFillObject;

static void patfill_dealloc(PyPatternFillObject *self) {
    free(self->fill_type);
    free(self->fg_color);
    free(self->bg_color);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *patfill_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    PyPatternFillObject *self = (PyPatternFillObject *)type->tp_alloc(type, 0);
    if (self) {
        self->fill_type = strdup("none");
        self->fg_color = NULL;
        self->bg_color = NULL;
    }
    return (PyObject *)self;
}

static int patfill_init(PyPatternFillObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"fill_type","fgColor","bgColor", NULL};
    const char *fill_type = "none";
    PyObject *fg = Py_None, *bg = Py_None;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|sOO", kwlist, &fill_type, &fg, &bg))
        return -1;

    free(self->fill_type);
    self->fill_type = strdup(fill_type);

    free(self->fg_color);
    self->fg_color = NULL;
    if (fg != Py_None) {
        if (!PyUnicode_Check(fg)) { PyErr_SetString(PyExc_TypeError, "fgColor must be a string or None"); return -1; }
        const char *s = PyUnicode_AsUTF8(fg);
        if (!s) return -1;
        self->fg_color = strdup(s);
        if (!self->fg_color) { PyErr_NoMemory(); return -1; }
    }

    free(self->bg_color);
    self->bg_color = NULL;
    if (bg != Py_None) {
        if (!PyUnicode_Check(bg)) { PyErr_SetString(PyExc_TypeError, "bgColor must be a string or None"); return -1; }
        const char *s = PyUnicode_AsUTF8(bg);
        if (!s) return -1;
        self->bg_color = strdup(s);
        if (!self->bg_color) { PyErr_NoMemory(); return -1; }
    }
    return 0;
}

static PyObject *patfill_get_fill_type(PyPatternFillObject *self, void *Py_UNUSED(x)) {
    if (!self->fill_type) return PyUnicode_FromString("none");
    return PyUnicode_FromString(self->fill_type);
}
static PyObject *patfill_get_fg(PyPatternFillObject *self, void *Py_UNUSED(x)) {
    if (!self->fg_color) Py_RETURN_NONE;
    return PyUnicode_FromString(self->fg_color);
}
static PyObject *patfill_get_bg(PyPatternFillObject *self, void *Py_UNUSED(x)) {
    if (!self->bg_color) Py_RETURN_NONE;
    return PyUnicode_FromString(self->bg_color);
}

static PyGetSetDef patfill_getset[] = {
    {"fill_type", (getter)patfill_get_fill_type, NULL, "Fill type", NULL},
    {"fgColor",   (getter)patfill_get_fg,        NULL, "Foreground color ARGB hex", NULL},
    {"bgColor",   (getter)patfill_get_bg,        NULL, "Background color ARGB hex", NULL},
    {NULL}
};

static PyTypeObject PyPatternFillType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.PatternFill",
    .tp_basicsize = sizeof(PyPatternFillObject),
    .tp_dealloc   = (destructor)patfill_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_new       = patfill_new,
    .tp_init      = (initproc)patfill_init,
    .tp_getset    = patfill_getset,
};

/* ---- PySideObject ---- */

typedef struct {
    PyObject_HEAD
    char *style;
    char *color;
} PySideObject;

static void side_dealloc(PySideObject *self) {
    free(self->style);
    free(self->color);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *side_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    PySideObject *self = (PySideObject *)type->tp_alloc(type, 0);
    if (self) {
        self->style = NULL;
        self->color = NULL;
    }
    return (PyObject *)self;
}

static int side_init(PySideObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"style","color", NULL};
    PyObject *style_obj = Py_None, *color_obj = Py_None;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|OO", kwlist, &style_obj, &color_obj))
        return -1;

    free(self->style);
    self->style = NULL;
    if (style_obj != Py_None) {
        if (!PyUnicode_Check(style_obj)) { PyErr_SetString(PyExc_TypeError, "style must be a string or None"); return -1; }
        const char *s = PyUnicode_AsUTF8(style_obj);
        if (!s) return -1;
        self->style = strdup(s);
        if (!self->style) { PyErr_NoMemory(); return -1; }
    }

    free(self->color);
    self->color = NULL;
    if (color_obj != Py_None) {
        if (!PyUnicode_Check(color_obj)) { PyErr_SetString(PyExc_TypeError, "color must be a string or None"); return -1; }
        const char *s = PyUnicode_AsUTF8(color_obj);
        if (!s) return -1;
        self->color = strdup(s);
        if (!self->color) { PyErr_NoMemory(); return -1; }
    }
    return 0;
}

static PyObject *side_get_style(PySideObject *self, void *Py_UNUSED(x)) {
    if (!self->style) Py_RETURN_NONE;
    return PyUnicode_FromString(self->style);
}
static PyObject *side_get_color(PySideObject *self, void *Py_UNUSED(x)) {
    if (!self->color) Py_RETURN_NONE;
    return PyUnicode_FromString(self->color);
}

static PyGetSetDef side_getset[] = {
    {"style", (getter)side_get_style, NULL, "Border style", NULL},
    {"color", (getter)side_get_color, NULL, "Color ARGB hex", NULL},
    {NULL}
};

static PyTypeObject PySideType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Side",
    .tp_basicsize = sizeof(PySideObject),
    .tp_dealloc   = (destructor)side_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_new       = side_new,
    .tp_init      = (initproc)side_init,
    .tp_getset    = side_getset,
};

/* ---- PyBorderObject ---- */

typedef struct {
    PyObject_HEAD
    PyObject *left;
    PyObject *right;
    PyObject *top;
    PyObject *bottom;
} PyBorderObject;

static int border_traverse(PyBorderObject *self, visitproc visit, void *arg) {
    Py_VISIT(self->left);
    Py_VISIT(self->right);
    Py_VISIT(self->top);
    Py_VISIT(self->bottom);
    return 0;
}

static int border_clear(PyBorderObject *self) {
    Py_CLEAR(self->left);
    Py_CLEAR(self->right);
    Py_CLEAR(self->top);
    Py_CLEAR(self->bottom);
    return 0;
}

static void border_dealloc(PyBorderObject *self) {
    PyObject_GC_UnTrack(self);
    border_clear(self);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *border_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    PyBorderObject *self = (PyBorderObject *)type->tp_alloc(type, 0);
    if (self) {
        Py_INCREF(Py_None);
        self->left = Py_None;
        Py_INCREF(Py_None);
        self->right = Py_None;
        Py_INCREF(Py_None);
        self->top = Py_None;
        Py_INCREF(Py_None);
        self->bottom = Py_None;
    }
    return (PyObject *)self;
}

static int border_init(PyBorderObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"left","right","top","bottom", NULL};
    PyObject *left = Py_None, *right = Py_None, *top = Py_None, *bottom = Py_None;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|OOOO", kwlist,
                                     &left, &right, &top, &bottom))
        return -1;

    /* Validate each is None or a PySideObject */
    if (left != Py_None && !PyObject_TypeCheck(left, &PySideType)) {
        PyErr_SetString(PyExc_TypeError, "left must be a Side or None"); return -1; }
    if (right != Py_None && !PyObject_TypeCheck(right, &PySideType)) {
        PyErr_SetString(PyExc_TypeError, "right must be a Side or None"); return -1; }
    if (top != Py_None && !PyObject_TypeCheck(top, &PySideType)) {
        PyErr_SetString(PyExc_TypeError, "top must be a Side or None"); return -1; }
    if (bottom != Py_None && !PyObject_TypeCheck(bottom, &PySideType)) {
        PyErr_SetString(PyExc_TypeError, "bottom must be a Side or None"); return -1; }

    Py_INCREF(left);   Py_DECREF(self->left);   self->left   = left;
    Py_INCREF(right);  Py_DECREF(self->right);  self->right  = right;
    Py_INCREF(top);    Py_DECREF(self->top);    self->top    = top;
    Py_INCREF(bottom); Py_DECREF(self->bottom); self->bottom = bottom;
    return 0;
}

static PyObject *border_get_left(PyBorderObject *self, void *Py_UNUSED(x)) {
    Py_INCREF(self->left); return self->left;
}
static PyObject *border_get_right(PyBorderObject *self, void *Py_UNUSED(x)) {
    Py_INCREF(self->right); return self->right;
}
static PyObject *border_get_top(PyBorderObject *self, void *Py_UNUSED(x)) {
    Py_INCREF(self->top); return self->top;
}
static PyObject *border_get_bottom(PyBorderObject *self, void *Py_UNUSED(x)) {
    Py_INCREF(self->bottom); return self->bottom;
}

static PyGetSetDef border_getset[] = {
    {"left",   (getter)border_get_left,   NULL, "Left side", NULL},
    {"right",  (getter)border_get_right,  NULL, "Right side", NULL},
    {"top",    (getter)border_get_top,    NULL, "Top side", NULL},
    {"bottom", (getter)border_get_bottom, NULL, "Bottom side", NULL},
    {NULL}
};

static PyTypeObject PyBorderType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Border",
    .tp_basicsize = sizeof(PyBorderObject),
    .tp_dealloc   = (destructor)border_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT | Py_TPFLAGS_HAVE_GC,
    .tp_new       = border_new,
    .tp_init      = (initproc)border_init,
    .tp_getset    = border_getset,
    .tp_traverse  = (traverseproc)border_traverse,
    .tp_clear     = (inquiry)border_clear,
};

/* ---- PyAlignmentObject ---- */

typedef struct {
    PyObject_HEAD
    char    *horizontal;
    char    *vertical;
    int32_t  indent;
    int32_t  text_rotation;
    uint8_t  wrap_text;
    uint8_t  shrink_to_fit;
} PyAlignmentObject;

static void align_dealloc(PyAlignmentObject *self) {
    free(self->horizontal);
    free(self->vertical);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *align_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    PyAlignmentObject *self = (PyAlignmentObject *)type->tp_alloc(type, 0);
    if (self) {
        self->horizontal = NULL;
        self->vertical = NULL;
        self->indent = 0;
        self->text_rotation = 0;
        self->wrap_text = 0;
        self->shrink_to_fit = 0;
    }
    return (PyObject *)self;
}

static int align_init(PyAlignmentObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {"horizontal","vertical","wrap_text","indent",
                              "text_rotation","shrink_to_fit", NULL};
    PyObject *horiz = Py_None, *vert = Py_None;
    int wrap_text = 0, indent = 0, text_rotation = 0, shrink_to_fit = 0;
    if (!PyArg_ParseTupleAndKeywords(args, kw, "|OOiiii", kwlist,
                                     &horiz, &vert, &wrap_text, &indent,
                                     &text_rotation, &shrink_to_fit))
        return -1;

    free(self->horizontal);
    self->horizontal = NULL;
    if (horiz != Py_None) {
        if (!PyUnicode_Check(horiz)) { PyErr_SetString(PyExc_TypeError, "horizontal must be str or None"); return -1; }
        const char *s = PyUnicode_AsUTF8(horiz);
        if (!s) return -1;
        self->horizontal = strdup(s);
        if (!self->horizontal) { PyErr_NoMemory(); return -1; }
    }

    free(self->vertical);
    self->vertical = NULL;
    if (vert != Py_None) {
        if (!PyUnicode_Check(vert)) { PyErr_SetString(PyExc_TypeError, "vertical must be str or None"); return -1; }
        const char *s = PyUnicode_AsUTF8(vert);
        if (!s) return -1;
        self->vertical = strdup(s);
        if (!self->vertical) { PyErr_NoMemory(); return -1; }
    }

    self->wrap_text = (uint8_t)(wrap_text ? 1 : 0);
    self->indent = (int32_t)indent;
    self->text_rotation = (int32_t)text_rotation;
    self->shrink_to_fit = (uint8_t)(shrink_to_fit ? 1 : 0);
    return 0;
}

static PyObject *align_get_horizontal(PyAlignmentObject *self, void *Py_UNUSED(x)) {
    if (!self->horizontal) Py_RETURN_NONE;
    return PyUnicode_FromString(self->horizontal);
}
static PyObject *align_get_vertical(PyAlignmentObject *self, void *Py_UNUSED(x)) {
    if (!self->vertical) Py_RETURN_NONE;
    return PyUnicode_FromString(self->vertical);
}
static PyObject *align_get_wrap_text(PyAlignmentObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->wrap_text);
}
static PyObject *align_get_indent(PyAlignmentObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromLong(self->indent);
}
static PyObject *align_get_text_rotation(PyAlignmentObject *self, void *Py_UNUSED(x)) {
    return PyLong_FromLong(self->text_rotation);
}
static PyObject *align_get_shrink_to_fit(PyAlignmentObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->shrink_to_fit);
}

static PyGetSetDef align_getset[] = {
    {"horizontal",    (getter)align_get_horizontal,    NULL, "Horizontal alignment", NULL},
    {"vertical",      (getter)align_get_vertical,      NULL, "Vertical alignment", NULL},
    {"wrap_text",     (getter)align_get_wrap_text,     NULL, "Wrap text", NULL},
    {"indent",        (getter)align_get_indent,        NULL, "Indent", NULL},
    {"text_rotation", (getter)align_get_text_rotation, NULL, "Text rotation", NULL},
    {"shrink_to_fit", (getter)align_get_shrink_to_fit, NULL, "Shrink to fit", NULL},
    {NULL}
};

static PyTypeObject PyAlignmentType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.Alignment",
    .tp_basicsize = sizeof(PyAlignmentObject),
    .tp_dealloc   = (destructor)align_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_new       = align_new,
    .tp_init      = (initproc)align_init,
    .tp_getset    = align_getset,
};

/* ---- PyXlDataValidationObject ---- */

typedef struct {
    PyObject_HEAD
    char *dv_type;
    char *dv_operator;
    char *formula1;
    char *formula2;
    char *sqref;
    char *error_message;
    char *error_title;
    char *error_style;
    char *prompt_message;
    char *prompt_title;
    int   allow_blank;
    int   show_drop_down;
    int   show_error_message;
    int   show_input_message;
} PyXlDataValidationObject;

static void dv_dealloc(PyXlDataValidationObject *self) {
    free(self->dv_type);
    free(self->dv_operator);
    free(self->formula1);
    free(self->formula2);
    free(self->sqref);
    free(self->error_message);
    free(self->error_title);
    free(self->error_style);
    free(self->prompt_message);
    free(self->prompt_title);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyObject *dv_new(PyTypeObject *type, PyObject *args, PyObject *kw) {
    (void)args; (void)kw;
    PyXlDataValidationObject *self = (PyXlDataValidationObject *)type->tp_alloc(type, 0);
    if (self) {
        self->dv_type        = NULL;
        self->dv_operator    = NULL;
        self->formula1       = NULL;
        self->formula2       = NULL;
        self->sqref          = NULL;
        self->error_message  = NULL;
        self->error_title    = NULL;
        self->error_style    = NULL;
        self->prompt_message = NULL;
        self->prompt_title   = NULL;
        self->allow_blank        = 1;  /* default True */
        self->show_drop_down     = 0;
        self->show_error_message = 0;
        self->show_input_message = 0;
    }
    return (PyObject *)self;
}

static int dv_init(PyXlDataValidationObject *self, PyObject *args, PyObject *kw) {
    static char *kwlist[] = {
        "type", "formula1", "formula2", "sqref",
        "allow_blank", "show_drop_down",
        "error_title", "error_message", "error_style",
        "prompt_title", "prompt_message",
        "show_error_message", "show_input_message",
        "operator",
        NULL
    };
    const char *dv_type    = NULL;
    const char *formula1   = NULL;
    const char *formula2   = NULL;
    const char *sqref      = NULL;
    const char *error_title   = NULL;
    const char *error_message = NULL;
    const char *error_style   = NULL;
    const char *prompt_title   = NULL;
    const char *prompt_message = NULL;
    const char *dv_operator = NULL;
    int allow_blank        = 1;
    int show_drop_down     = 0;
    int show_error_message = 0;
    int show_input_message = 0;

    if (!PyArg_ParseTupleAndKeywords(args, kw,
            "|zzzziizzzzziiiz", kwlist,
            &dv_type, &formula1, &formula2, &sqref,
            &allow_blank, &show_drop_down,
            &error_title, &error_message, &error_style,
            &prompt_title, &prompt_message,
            &show_error_message, &show_input_message,
            &dv_operator))
        return -1;

    free(self->dv_type);        self->dv_type        = dv_type        ? strdup(dv_type)        : NULL;
    free(self->dv_operator);    self->dv_operator    = dv_operator    ? strdup(dv_operator)    : NULL;
    free(self->formula1);       self->formula1       = formula1       ? strdup(formula1)       : NULL;
    free(self->formula2);       self->formula2       = formula2       ? strdup(formula2)       : NULL;
    free(self->sqref);          self->sqref          = sqref          ? strdup(sqref)          : NULL;
    free(self->error_title);    self->error_title    = error_title    ? strdup(error_title)    : NULL;
    free(self->error_message);  self->error_message  = error_message  ? strdup(error_message)  : NULL;
    free(self->error_style);    self->error_style    = error_style    ? strdup(error_style)    : NULL;
    free(self->prompt_title);   self->prompt_title   = prompt_title   ? strdup(prompt_title)   : NULL;
    free(self->prompt_message); self->prompt_message = prompt_message ? strdup(prompt_message) : NULL;
    self->allow_blank        = allow_blank;
    self->show_drop_down     = show_drop_down;
    self->show_error_message = show_error_message;
    self->show_input_message = show_input_message;
    return 0;
}

static PyObject *dv_get_type(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->dv_type) Py_RETURN_NONE;
    return PyUnicode_FromString(self->dv_type);
}
static PyObject *dv_get_operator(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->dv_operator) Py_RETURN_NONE;
    return PyUnicode_FromString(self->dv_operator);
}
static PyObject *dv_get_formula1(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->formula1) Py_RETURN_NONE;
    return PyUnicode_FromString(self->formula1);
}
static PyObject *dv_get_formula2(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->formula2) Py_RETURN_NONE;
    return PyUnicode_FromString(self->formula2);
}
static PyObject *dv_get_sqref(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->sqref) Py_RETURN_NONE;
    return PyUnicode_FromString(self->sqref);
}
static PyObject *dv_get_error_message(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->error_message) Py_RETURN_NONE;
    return PyUnicode_FromString(self->error_message);
}
static PyObject *dv_get_error_title(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->error_title) Py_RETURN_NONE;
    return PyUnicode_FromString(self->error_title);
}
static PyObject *dv_get_error_style(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->error_style) Py_RETURN_NONE;
    return PyUnicode_FromString(self->error_style);
}
static PyObject *dv_get_prompt_message(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->prompt_message) Py_RETURN_NONE;
    return PyUnicode_FromString(self->prompt_message);
}
static PyObject *dv_get_prompt_title(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    if (!self->prompt_title) Py_RETURN_NONE;
    return PyUnicode_FromString(self->prompt_title);
}
static PyObject *dv_get_allow_blank(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->allow_blank);
}
static PyObject *dv_get_show_drop_down(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->show_drop_down);
}
static PyObject *dv_get_show_error_message(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->show_error_message);
}
static PyObject *dv_get_show_input_message(PyXlDataValidationObject *self, void *Py_UNUSED(x)) {
    return PyBool_FromLong(self->show_input_message);
}

static PyGetSetDef dv_getset[] = {
    {"type",               (getter)dv_get_type,               NULL, "Validation type", NULL},
    {"operator",           (getter)dv_get_operator,           NULL, "Operator", NULL},
    {"formula1",           (getter)dv_get_formula1,           NULL, "Formula 1", NULL},
    {"formula2",           (getter)dv_get_formula2,           NULL, "Formula 2", NULL},
    {"sqref",              (getter)dv_get_sqref,              NULL, "Cell range", NULL},
    {"error_message",      (getter)dv_get_error_message,      NULL, "Error message", NULL},
    {"error_title",        (getter)dv_get_error_title,        NULL, "Error title", NULL},
    {"error_style",        (getter)dv_get_error_style,        NULL, "Error style", NULL},
    {"prompt_message",     (getter)dv_get_prompt_message,     NULL, "Prompt message", NULL},
    {"prompt_title",       (getter)dv_get_prompt_title,       NULL, "Prompt title", NULL},
    {"allow_blank",        (getter)dv_get_allow_blank,        NULL, "Allow blank", NULL},
    {"show_drop_down",     (getter)dv_get_show_drop_down,     NULL, "Show drop-down (False=show)", NULL},
    {"show_error_message", (getter)dv_get_show_error_message, NULL, "Show error message", NULL},
    {"show_input_message", (getter)dv_get_show_input_message, NULL, "Show input message", NULL},
    {NULL}
};

static PyTypeObject PyXlDataValidationType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name      = "_openexcel.DataValidation",
    .tp_basicsize = sizeof(PyXlDataValidationObject),
    .tp_dealloc   = (destructor)dv_dealloc,
    .tp_flags     = Py_TPFLAGS_DEFAULT,
    .tp_new       = dv_new,
    .tp_init      = (initproc)dv_init,
    .tp_getset    = dv_getset,
};

/* ========== Phase 3: Cell XF context helper ========== */

static void get_cell_xf_context(PyXlCellObject *self,
                                  uint16_t *out_xf_idx,
                                  const char **out_fmt_str,
                                  uint32_t *out_font_id,
                                  uint32_t *out_fill_id,
                                  uint32_t *out_border_id,
                                  OxlAlignDef *out_align) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    *out_xf_idx = c ? c->style_idx : 0;
    const OxlXfRecord *xf = oxl_styles_get_xf(&self->wb->styles, *out_xf_idx);
    *out_fmt_str = xf ? xf->fmt_str : NULL;
    *out_font_id = xf ? xf->font_id : 0;
    *out_fill_id = xf ? xf->fill_id : 0;
    *out_border_id = xf ? xf->border_id : 0;
    if (out_align) {
        if (xf) *out_align = xf->align;
        else memset(out_align, 0, sizeof(*out_align));
    }
}

/* Helper: ensure a cell slot exists (create stub EMPTY cell if needed) and set style_idx */
static int ensure_cell_with_style(PyXlCellObject *self, uint16_t xf_idx) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    if (c) {
        c->style_idx = xf_idx;
        return 0;
    }
    /* Create stub EMPTY cell to carry the style */
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
    /* Update row_count and col_count */
    if (self->row + 1 > self->ws->row_count) self->ws->row_count = self->row + 1;
    if ((uint32_t)self->col + 1 > self->ws->col_count) self->ws->col_count = (uint32_t)self->col + 1;
    return 0;
}

/* ========== Phase 3: Cell getters/setters for font/fill/border/alignment ========== */

/* Helper: build a PySideObject from an OxlBorderSide */
static PyObject *make_side_from_c(const OxlBorderSide *side) {
    PySideObject *obj = (PySideObject *)PySideType.tp_alloc(&PySideType, 0);
    if (!obj) return NULL;
    obj->style = (side && side->style) ? strdup(side->style) : NULL;
    obj->color = (side && side->has_color) ? NULL : NULL;  /* color below */
    if (side && side->has_color) {
        char buf[9];
        snprintf(buf, sizeof(buf), "%08X", side->color_rgb);
        obj->color = strdup(buf);
    }
    return (PyObject *)obj;
}

static PyObject *cell_get_font(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    uint16_t xf_idx = c ? c->style_idx : 0;
    const OxlXfRecord *xf = oxl_styles_get_xf(&self->wb->styles, xf_idx);
    const OxlFontDef *font = oxl_styles_get_font(&self->wb->styles, xf ? xf->font_id : 0);

    PyFontObject *obj = (PyFontObject *)PyFontType.tp_alloc(&PyFontType, 0);
    if (!obj) return NULL;

    if (font) {
        obj->name = font->name ? strdup(font->name) : NULL;
        obj->size = font->size > 0 ? (double)font->size : 11.0;
        obj->bold = font->bold;
        obj->italic = font->italic;
        obj->underline = font->underline;
        if (font->color_rgb != 0) {
            char buf[9];
            snprintf(buf, sizeof(buf), "%08X", font->color_rgb);
            obj->color = strdup(buf);
        } else {
            obj->color = NULL;
        }
    } else {
        obj->name = strdup("Calibri");
        obj->size = 11.0;
        obj->bold = 0;
        obj->italic = 0;
        obj->underline = 0;
        obj->color = NULL;
    }
    return (PyObject *)obj;
}

static int cell_set_font(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value) { PyErr_SetString(PyExc_TypeError, "Cannot delete font"); return -1; }

    uint16_t xf_idx;
    const char *fmt_str;
    uint32_t font_id, fill_id, border_id;
    OxlAlignDef align;
    get_cell_xf_context(self, &xf_idx, &fmt_str, &font_id, &fill_id, &border_id, &align);

    oxl_styles_init_write_defaults(&self->wb->styles);

    if (value == Py_None) {
        font_id = 0;
    } else if (PyObject_TypeCheck(value, &PyFontType)) {
        PyFontObject *fo = (PyFontObject *)value;
        OxlFontDef fd;
        memset(&fd, 0, sizeof(fd));
        fd.name = fo->name;
        fd.size = (float)fo->size;
        fd.bold = fo->bold;
        fd.italic = fo->italic;
        fd.underline = fo->underline;
        fd.color_rgb = fo->color ? parse_color_str(fo->color) : 0;
        font_id = oxl_styles_get_or_add_font(&self->wb->styles, &fd);
    } else {
        PyErr_SetString(PyExc_TypeError, "font must be a Font object or None");
        return -1;
    }

    uint16_t new_xf = oxl_styles_get_or_add_xf_full(&self->wb->styles, fmt_str,
                                                      font_id, fill_id, border_id, &align);
    return ensure_cell_with_style(self, new_xf);
}

static PyObject *cell_get_fill(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    uint16_t xf_idx = c ? c->style_idx : 0;
    const OxlXfRecord *xf = oxl_styles_get_xf(&self->wb->styles, xf_idx);
    const OxlFillDef *fill = oxl_styles_get_fill(&self->wb->styles, xf ? xf->fill_id : 0);

    PyPatternFillObject *obj = (PyPatternFillObject *)PyPatternFillType.tp_alloc(&PyPatternFillType, 0);
    if (!obj) return NULL;

    if (fill) {
        obj->fill_type = fill->pattern_type ? strdup(fill->pattern_type) : strdup("none");
        obj->fg_color = (fill->fg_has_color && fill->fg_rgb != 0) ? NULL : NULL;
        if (fill->fg_has_color && fill->fg_rgb != 0) {
            char buf[9];
            snprintf(buf, sizeof(buf), "%08X", fill->fg_rgb);
            obj->fg_color = strdup(buf);
        }
        obj->bg_color = NULL;
        if (fill->bg_has_color && fill->bg_rgb != 0) {
            char buf[9];
            snprintf(buf, sizeof(buf), "%08X", fill->bg_rgb);
            obj->bg_color = strdup(buf);
        }
    } else {
        obj->fill_type = strdup("none");
        obj->fg_color = NULL;
        obj->bg_color = NULL;
    }
    return (PyObject *)obj;
}

static int cell_set_fill(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value) { PyErr_SetString(PyExc_TypeError, "Cannot delete fill"); return -1; }

    uint16_t xf_idx;
    const char *fmt_str;
    uint32_t font_id, fill_id, border_id;
    OxlAlignDef align;
    get_cell_xf_context(self, &xf_idx, &fmt_str, &font_id, &fill_id, &border_id, &align);

    oxl_styles_init_write_defaults(&self->wb->styles);

    if (value == Py_None) {
        fill_id = 0;
    } else if (PyObject_TypeCheck(value, &PyPatternFillType)) {
        PyPatternFillObject *fo = (PyPatternFillObject *)value;
        OxlFillDef fd;
        memset(&fd, 0, sizeof(fd));
        fd.pattern_type = fo->fill_type;
        if (fo->fg_color) {
            fd.fg_rgb = parse_color_str(fo->fg_color);
            fd.fg_has_color = (fd.fg_rgb != 0) ? 1 : 0;
        }
        if (fo->bg_color) {
            fd.bg_rgb = parse_color_str(fo->bg_color);
            fd.bg_has_color = (fd.bg_rgb != 0) ? 1 : 0;
        }
        fill_id = oxl_styles_get_or_add_fill(&self->wb->styles, &fd);
    } else {
        PyErr_SetString(PyExc_TypeError, "fill must be a PatternFill object or None");
        return -1;
    }

    uint16_t new_xf = oxl_styles_get_or_add_xf_full(&self->wb->styles, fmt_str,
                                                      font_id, fill_id, border_id, &align);
    return ensure_cell_with_style(self, new_xf);
}

static PyObject *cell_get_border(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    uint16_t xf_idx = c ? c->style_idx : 0;
    const OxlXfRecord *xf = oxl_styles_get_xf(&self->wb->styles, xf_idx);
    const OxlBorderDef *border = oxl_styles_get_border(&self->wb->styles, xf ? xf->border_id : 0);

    PyBorderObject *obj = (PyBorderObject *)PyBorderType.tp_alloc(&PyBorderType, 0);
    if (!obj) return NULL;
    /* Initialize all to None */
    Py_INCREF(Py_None); obj->left   = Py_None;
    Py_INCREF(Py_None); obj->right  = Py_None;
    Py_INCREF(Py_None); obj->top    = Py_None;
    Py_INCREF(Py_None); obj->bottom = Py_None;

    if (border) {
        if (border->left.style) {
            PyObject *s = make_side_from_c(&border->left);
            if (!s) { Py_DECREF(obj); return NULL; }
            Py_DECREF(obj->left); obj->left = s;
        }
        if (border->right.style) {
            PyObject *s = make_side_from_c(&border->right);
            if (!s) { Py_DECREF(obj); return NULL; }
            Py_DECREF(obj->right); obj->right = s;
        }
        if (border->top.style) {
            PyObject *s = make_side_from_c(&border->top);
            if (!s) { Py_DECREF(obj); return NULL; }
            Py_DECREF(obj->top); obj->top = s;
        }
        if (border->bottom.style) {
            PyObject *s = make_side_from_c(&border->bottom);
            if (!s) { Py_DECREF(obj); return NULL; }
            Py_DECREF(obj->bottom); obj->bottom = s;
        }
    }
    return (PyObject *)obj;
}

static int cell_set_border(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value) { PyErr_SetString(PyExc_TypeError, "Cannot delete border"); return -1; }

    uint16_t xf_idx;
    const char *fmt_str;
    uint32_t font_id, fill_id, border_id;
    OxlAlignDef align;
    get_cell_xf_context(self, &xf_idx, &fmt_str, &font_id, &fill_id, &border_id, &align);

    oxl_styles_init_write_defaults(&self->wb->styles);

    if (value == Py_None) {
        border_id = 0;
    } else if (PyObject_TypeCheck(value, &PyBorderType)) {
        PyBorderObject *bo = (PyBorderObject *)value;
        OxlBorderDef bd;
        memset(&bd, 0, sizeof(bd));

        if (bo->left != Py_None) {
            PySideObject *s = (PySideObject *)bo->left;
            bd.left.style = s->style;
            if (s->color) { bd.left.color_rgb = parse_color_str(s->color); bd.left.has_color = 1; }
        }
        if (bo->right != Py_None) {
            PySideObject *s = (PySideObject *)bo->right;
            bd.right.style = s->style;
            if (s->color) { bd.right.color_rgb = parse_color_str(s->color); bd.right.has_color = 1; }
        }
        if (bo->top != Py_None) {
            PySideObject *s = (PySideObject *)bo->top;
            bd.top.style = s->style;
            if (s->color) { bd.top.color_rgb = parse_color_str(s->color); bd.top.has_color = 1; }
        }
        if (bo->bottom != Py_None) {
            PySideObject *s = (PySideObject *)bo->bottom;
            bd.bottom.style = s->style;
            if (s->color) { bd.bottom.color_rgb = parse_color_str(s->color); bd.bottom.has_color = 1; }
        }
        border_id = oxl_styles_get_or_add_border(&self->wb->styles, &bd);
    } else {
        PyErr_SetString(PyExc_TypeError, "border must be a Border object or None");
        return -1;
    }

    uint16_t new_xf = oxl_styles_get_or_add_xf_full(&self->wb->styles, fmt_str,
                                                      font_id, fill_id, border_id, &align);
    return ensure_cell_with_style(self, new_xf);
}

static PyObject *cell_get_alignment(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    uint16_t xf_idx = c ? c->style_idx : 0;
    const OxlXfRecord *xf = oxl_styles_get_xf(&self->wb->styles, xf_idx);

    PyAlignmentObject *obj = (PyAlignmentObject *)PyAlignmentType.tp_alloc(&PyAlignmentType, 0);
    if (!obj) return NULL;
    obj->horizontal = NULL;
    obj->vertical = NULL;
    obj->indent = 0;
    obj->text_rotation = 0;
    obj->wrap_text = 0;
    obj->shrink_to_fit = 0;

    if (xf && xf->apply_alignment) {
        const OxlAlignDef *a = &xf->align;
        obj->horizontal    = a->horizontal   ? strdup(a->horizontal)  : NULL;
        obj->vertical      = a->vertical     ? strdup(a->vertical)    : NULL;
        obj->indent        = a->indent;
        obj->text_rotation = a->text_rotation;
        obj->wrap_text     = a->wrap_text;
        obj->shrink_to_fit = a->shrink_to_fit;
    }
    return (PyObject *)obj;
}

static int cell_set_alignment(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value) { PyErr_SetString(PyExc_TypeError, "Cannot delete alignment"); return -1; }

    uint16_t xf_idx;
    const char *fmt_str;
    uint32_t font_id, fill_id, border_id;
    OxlAlignDef old_align;
    get_cell_xf_context(self, &xf_idx, &fmt_str, &font_id, &fill_id, &border_id, &old_align);
    (void)old_align;

    oxl_styles_init_write_defaults(&self->wb->styles);

    OxlAlignDef align_def;
    memset(&align_def, 0, sizeof(align_def));

    if (value == Py_None) {
        /* align_def stays zero = no alignment */
    } else if (PyObject_TypeCheck(value, &PyAlignmentType)) {
        PyAlignmentObject *ao = (PyAlignmentObject *)value;
        align_def.horizontal    = ao->horizontal;
        align_def.vertical      = ao->vertical;
        align_def.indent        = ao->indent;
        align_def.text_rotation = ao->text_rotation;
        align_def.wrap_text     = ao->wrap_text;
        align_def.shrink_to_fit = ao->shrink_to_fit;
    } else {
        PyErr_SetString(PyExc_TypeError, "alignment must be an Alignment object or None");
        return -1;
    }

    uint16_t new_xf = oxl_styles_get_or_add_xf_full(&self->wb->styles, fmt_str,
                                                      font_id, fill_id, border_id, &align_def);
    return ensure_cell_with_style(self, new_xf);
}

static PyObject *cell_get_hyperlink(PyXlCellObject *self, void *Py_UNUSED(x)) {
    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    if (!c || !c->hyperlink) Py_RETURN_NONE;
    return PyUnicode_FromString(c->hyperlink);
}

static int cell_set_hyperlink(PyXlCellObject *self, PyObject *value, void *Py_UNUSED(x)) {
    if (!value || value == Py_None) {
        OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
        if (c) { free(c->hyperlink); c->hyperlink = NULL; }
        return 0;
    }
    if (!PyUnicode_Check(value)) {
        PyErr_SetString(PyExc_TypeError, "hyperlink must be a string or None");
        return -1;
    }
    const char *url = PyUnicode_AsUTF8(value);
    if (!url) return -1;

    OxlCell *c = ws_find_cell(self->ws, self->row, self->col);
    if (c) {
        free(c->hyperlink);
        c->hyperlink = strdup(url);
    } else {
        /* Create a stub EMPTY cell to carry the hyperlink */
        OxlCell newc;
        memset(&newc, 0, sizeof(newc));
        newc.row = self->row;
        newc.col = self->col;
        newc.hyperlink = strdup(url);
        newc.type = OXL_CELL_EMPTY;
        uint32_t pos = ws_find_pos(self->ws, self->row, self->col);
        if (self->ws->cell_count >= self->ws->cell_capacity) {
            uint32_t new_cap = self->ws->cell_capacity == 0 ? 8 : self->ws->cell_capacity * 2;
            OxlCell *nc = (OxlCell *)realloc(self->ws->cells, new_cap * sizeof(OxlCell));
            if (!nc) { free(newc.hyperlink); PyErr_NoMemory(); return -1; }
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
    {"font",          (getter)cell_get_font,            (setter)cell_set_font,           "Cell font", NULL},
    {"fill",          (getter)cell_get_fill,            (setter)cell_set_fill,           "Cell fill", NULL},
    {"border",        (getter)cell_get_border,          (setter)cell_set_border,         "Cell border", NULL},
    {"alignment",     (getter)cell_get_alignment,       (setter)cell_set_alignment,      "Cell alignment", NULL},
    {"hyperlink",     (getter)cell_get_hyperlink,       (setter)cell_set_hyperlink,      "Cell hyperlink URL", NULL},
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

/* ── Phase 13: Data Validations ──────────────────────────────────────────── */

static PyObject *ws_add_data_validation(PyObject *self, PyObject *args) {
    PyObject *dv_obj;
    if (!PyArg_ParseTuple(args, "O!", &PyXlDataValidationType, &dv_obj))
        return NULL;
    PyXlDataValidationObject *dv = (PyXlDataValidationObject *)dv_obj;
    OxlWorksheet *ws = ((PyWorksheetObject *)self)->ws;
    OxlDataValidation cdv;
    memset(&cdv, 0, sizeof(cdv));
    cdv.type               = dv->dv_type;
    cdv.dv_operator        = dv->dv_operator;
    cdv.formula1           = dv->formula1;
    cdv.formula2           = dv->formula2;
    cdv.sqref              = dv->sqref;
    cdv.error_message      = dv->error_message;
    cdv.error_title        = dv->error_title;
    cdv.error_style        = dv->error_style;
    cdv.prompt_message     = dv->prompt_message;
    cdv.prompt_title       = dv->prompt_title;
    cdv.allow_blank        = (uint8_t)dv->allow_blank;
    cdv.show_drop_down     = (uint8_t)dv->show_drop_down;
    cdv.show_error_message = (uint8_t)dv->show_error_message;
    cdv.show_input_message = (uint8_t)dv->show_input_message;
    if (oxl_worksheet_add_data_validation(ws, &cdv) < 0) {
        PyErr_NoMemory(); return NULL;
    }
    Py_RETURN_NONE;
}

static PyObject *ws_get_data_validations(PyObject *self, void *closure) {
    (void)closure;
    OxlWorksheet *ws = ((PyWorksheetObject *)self)->ws;
    PyObject *list = PyList_New((Py_ssize_t)ws->dv_count);
    if (!list) return NULL;
    for (uint32_t i = 0; i < ws->dv_count; i++) {
        const OxlDataValidation *dv = &ws->data_validations[i];
        PyXlDataValidationObject *obj = PyObject_New(PyXlDataValidationObject, &PyXlDataValidationType);
        if (!obj) { Py_DECREF(list); return NULL; }
        obj->dv_type        = dv->type           ? strdup(dv->type)           : NULL;
        obj->dv_operator    = dv->dv_operator     ? strdup(dv->dv_operator)     : NULL;
        obj->formula1       = dv->formula1        ? strdup(dv->formula1)        : NULL;
        obj->formula2       = dv->formula2        ? strdup(dv->formula2)        : NULL;
        obj->sqref          = dv->sqref           ? strdup(dv->sqref)           : NULL;
        obj->error_message  = dv->error_message   ? strdup(dv->error_message)   : NULL;
        obj->error_title    = dv->error_title     ? strdup(dv->error_title)     : NULL;
        obj->error_style    = dv->error_style     ? strdup(dv->error_style)     : NULL;
        obj->prompt_message = dv->prompt_message  ? strdup(dv->prompt_message)  : NULL;
        obj->prompt_title   = dv->prompt_title    ? strdup(dv->prompt_title)    : NULL;
        obj->allow_blank        = dv->allow_blank;
        obj->show_drop_down     = dv->show_drop_down;
        obj->show_error_message = dv->show_error_message;
        obj->show_input_message = dv->show_input_message;
        PyList_SET_ITEM(list, (Py_ssize_t)i, (PyObject *)obj);
    }
    return list;
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
    {"set_row_height",         (PyCFunction)(void(*)(void))worksheet_set_row_height,   METH_VARARGS|METH_KEYWORDS, "set_row_height(row, height, hidden=False)"},
    {"add_data_validation",    (PyCFunction)ws_add_data_validation,                    METH_VARARGS,               "add_data_validation(dv)"},
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
    {"auto_filter_ref",   (getter)worksheet_get_auto_filter_ref, (setter)worksheet_set_auto_filter_ref, "Auto-filter range",  NULL},
    {"data_validations",  (getter)ws_get_data_validations,       NULL,                                  "Data validations",   NULL},
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
    if (PyType_Ready(&PyXlCellType) < 0)      return NULL;
    if (PyType_Ready(&PyWorksheetType) < 0)   return NULL;
    if (PyType_Ready(&PyWorkbookType) < 0)    return NULL;
    if (PyType_Ready(&PyFontType) < 0)        return NULL;
    if (PyType_Ready(&PyPatternFillType) < 0) return NULL;
    if (PyType_Ready(&PySideType) < 0)        return NULL;
    if (PyType_Ready(&PyBorderType) < 0)      return NULL;
    if (PyType_Ready(&PyAlignmentType) < 0)          return NULL;
    if (PyType_Ready(&PyXlDataValidationType) < 0)   return NULL;

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

    Py_INCREF(&PyFontType);
    if (PyModule_AddObject(mod, "Font", (PyObject *)&PyFontType) < 0) {
        Py_DECREF(&PyFontType); Py_DECREF(mod); return NULL;
    }

    Py_INCREF(&PyPatternFillType);
    if (PyModule_AddObject(mod, "PatternFill", (PyObject *)&PyPatternFillType) < 0) {
        Py_DECREF(&PyPatternFillType); Py_DECREF(mod); return NULL;
    }

    Py_INCREF(&PySideType);
    if (PyModule_AddObject(mod, "Side", (PyObject *)&PySideType) < 0) {
        Py_DECREF(&PySideType); Py_DECREF(mod); return NULL;
    }

    Py_INCREF(&PyBorderType);
    if (PyModule_AddObject(mod, "Border", (PyObject *)&PyBorderType) < 0) {
        Py_DECREF(&PyBorderType); Py_DECREF(mod); return NULL;
    }

    Py_INCREF(&PyAlignmentType);
    if (PyModule_AddObject(mod, "Alignment", (PyObject *)&PyAlignmentType) < 0) {
        Py_DECREF(&PyAlignmentType); Py_DECREF(mod); return NULL;
    }

    Py_INCREF(&PyXlDataValidationType);
    if (PyModule_AddObject(mod, "DataValidation", (PyObject *)&PyXlDataValidationType) < 0) {
        Py_DECREF(&PyXlDataValidationType); Py_DECREF(mod); return NULL;
    }

    return mod;
}
