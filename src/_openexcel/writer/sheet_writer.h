#pragma once
#include "../worksheet.h"
#include "../workbook.h"
#include "xml_gen.h"

/* Emit sheetN.xml for the given worksheet into b. */
void oxl_write_sheet(OxlXmlBuf *b, const OxlWorksheet *ws, const OxlWorkbook *wb,
                     uint16_t date_xf_idx);
