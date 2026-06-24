#pragma once
#include <stddef.h>
#include "../workbook.h"

/* Parse xl/worksheets/sheetN.xml into ws, using wb for SST and styles lookups.
   Returns 0 on success, -1 on parse error. */
int oxl_parse_sheet(const char *buf, size_t len,
                    OxlWorksheet *ws, OxlWorkbook *wb);
