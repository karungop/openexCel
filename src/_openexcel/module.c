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
#include "string_table.h"
#include "reader/reader.h"
#include "writer/writer.h"

/* ========== Forward declarations ========== */
static PyTypeObject PyWorkbookType;
static PyTypeObject PyWorksheetType;
static PyTypeObject PyRowIteratorType;

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

/* ========== Helper: OxlCell → Python object ========== */

static PyObject *cell_to_python(const OxlCell *c, OxlWorkbook *wb) {
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

static void worksheet_dealloc(PyWorksheetObject *self) {
    Py_XDECREF(self->owner);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyMethodDef worksheet_methods[] = {
    {"__iter__",  (PyCFunction)worksheet_iter,      METH_NOARGS,   "Iterate rows"},
    {"iter_rows", (PyCFunction)(void(*)(void))worksheet_iter_rows, METH_VARARGS|METH_KEYWORDS, "iter_rows(min_row=0, max_row=-1, min_col=0, max_col=-1)"},
    {"append",    (PyCFunction)worksheet_append,    METH_O,        "Append a row"},
    {NULL, NULL}
};

static PyGetSetDef worksheet_getset[] = {
    {"max_row",    (getter)worksheet_get_max_row, NULL, "Number of rows", NULL},
    {"max_column", (getter)worksheet_get_max_col, NULL, "Number of columns", NULL},
    {"title",      (getter)worksheet_get_title,   NULL, "Sheet name", NULL},
    {NULL}
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

static PyMethodDef workbook_methods[] = {
    {"create_sheet", (PyCFunction)workbook_create_sheet, METH_VARARGS, "create_sheet(name='Sheet')"},
    {"save",         (PyCFunction)workbook_save,         METH_VARARGS, "save(path)"},
    {"close",        (PyCFunction)workbook_close,        METH_NOARGS,  "close()"},
    {"__enter__",    (PyCFunction)workbook_enter,        METH_NOARGS,  NULL},
    {"__exit__",     (PyCFunction)workbook_exit,         METH_VARARGS, NULL},
    {"__len__",      (PyCFunction)workbook_len,          METH_NOARGS,  NULL},
    {NULL, NULL}
};

static PyGetSetDef workbook_getset[] = {
    {"active", (getter)workbook_get_active, NULL, "First sheet", NULL},
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

/* ========== Module definition ========== */

static PyMethodDef module_methods[] = {
    {"load_workbook", m_load_workbook, METH_VARARGS, "load_workbook(path) -> Workbook"},
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
    if (PyType_Ready(&PyWorksheetType) < 0)  return NULL;
    if (PyType_Ready(&PyWorkbookType) < 0)   return NULL;

    PyObject *mod = PyModule_Create(&moduledef);
    if (!mod) return NULL;

    Py_INCREF(&PyWorkbookType);
    if (PyModule_AddObject(mod, "Workbook", (PyObject *)&PyWorkbookType) < 0) {
        Py_DECREF(&PyWorkbookType);
        Py_DECREF(mod);
        return NULL;
    }

    return mod;
}
