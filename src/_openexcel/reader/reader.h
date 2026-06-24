#pragma once
#include <stddef.h>
#include "../workbook.h"

/* Read an .xlsx file from disk. Returns a new OxlWorkbook* or NULL on error.
   Caller must call oxl_workbook_free() when done. */
OxlWorkbook *oxl_read_workbook(const char *path, char *err_buf, size_t err_cap);
