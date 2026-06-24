#pragma once
#include <stddef.h>
#include "../workbook.h"

/* Parse workbook.xml: extract date1904 flag and populate sheet name list.
   rel_path for each sheet is populated later by parsing workbook.xml.rels. */
int oxl_parse_workbook_xml(const char *buf, size_t len, OxlWorkbook *wb);

/* Parse xl/_rels/workbook.xml.rels: map rId → rel_path.
   Fills in ws->rel_path for sheets already added to wb. */
int oxl_parse_workbook_rels(const char *buf, size_t len, OxlWorkbook *wb);
