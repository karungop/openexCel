#pragma once
#include <stddef.h>
#include "../workbook.h"

/* Write wb as an .xlsx file to path.
   Returns 0 on success, -1 on error (err_buf filled if non-NULL). */
int oxl_write_workbook(OxlWorkbook *wb, const char *path,
                       char *err_buf, size_t err_cap);
