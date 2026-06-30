#pragma once
#include <stddef.h>
#include "../workbook.h"
#include "xml_sheet_rels.h"

/* Parse xl/worksheets/sheetN.xml into ws, using wb for SST/styles lookups.
   rels may be NULL if no rels file exists for this sheet.
   Returns 0 on success, -1 on parse error. */
int oxl_parse_sheet(const char *buf, size_t len,
                    OxlWorksheet *ws, OxlWorkbook *wb,
                    const OxlHyperlinkRels *rels);
