#include "sheet_writer.h"
#include "../cell.h"
#include <stdint.h>
#include <string.h>
#include <stdio.h>

static void emit_cell_ref(OxlXmlBuf *b, uint32_t row, uint16_t col) {
    oxl_xmlbuf_col_label(b, col);
    oxl_xmlbuf_uint(b, row + 1);  /* xlsx is 1-based */
}

void oxl_write_sheet(OxlXmlBuf *b, const OxlWorksheet *ws, const OxlWorkbook *wb,
                     uint16_t date_xf_idx) {
    (void)wb;
    oxl_xmlbuf_cstr(b,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<sheetData>");

    uint32_t i = 0;
    while (i < ws->cell_count) {
        uint32_t row = ws->cells[i].row;
        /* Open row */
        oxl_xmlbuf_cstr(b, "<row r=\"");
        oxl_xmlbuf_uint(b, row + 1);
        oxl_xmlbuf_cstr(b, "\">");

        /* Emit all cells in this row */
        while (i < ws->cell_count && ws->cells[i].row == row) {
            const OxlCell *c = &ws->cells[i];
            i++;

            if (c->type == OXL_CELL_EMPTY) continue;

            oxl_xmlbuf_cstr(b, "<c r=\"");
            emit_cell_ref(b, c->row, c->col);
            oxl_xmlbuf_raw(b, "\"", 1);

            switch (c->type) {
            case OXL_CELL_STRING:
            case OXL_CELL_INLINE_STR:
                oxl_xmlbuf_cstr(b, " t=\"s\"><v>");
                if (c->type == OXL_CELL_STRING) {
                    oxl_xmlbuf_uint(b, c->v.s_idx);
                } else {
                    /* inline str was already interned during write prep */
                    oxl_xmlbuf_uint(b, c->v.s_idx);
                }
                oxl_xmlbuf_cstr(b, "</v></c>");
                break;
            case OXL_CELL_BOOL:
                oxl_xmlbuf_cstr(b, " t=\"b\"><v>");
                oxl_xmlbuf_raw(b, c->v.b ? "1" : "0", 1);
                oxl_xmlbuf_cstr(b, "</v></c>");
                break;
            case OXL_CELL_FLOAT:
                oxl_xmlbuf_cstr(b, "><v>");
                oxl_xmlbuf_double(b, c->v.f);
                oxl_xmlbuf_cstr(b, "</v></c>");
                break;
            case OXL_CELL_DATE: {
                double serial = oxl_date_to_serial(c->v.dt, wb->date1904);
                oxl_xmlbuf_cstr(b, " s=\"");
                oxl_xmlbuf_uint(b, date_xf_idx);
                oxl_xmlbuf_cstr(b, "\"><v>");
                oxl_xmlbuf_double(b, serial);
                oxl_xmlbuf_cstr(b, "</v></c>");
                break;
            }
            case OXL_CELL_ERROR:
                oxl_xmlbuf_cstr(b, " t=\"e\"><v>");
                oxl_xmlbuf_text(b, c->v.s_inline ? c->v.s_inline : "#ERR!");
                oxl_xmlbuf_cstr(b, "</v></c>");
                break;
            default:
                break;
            }
        }
        oxl_xmlbuf_cstr(b, "</row>");
    }
    oxl_xmlbuf_cstr(b, "</sheetData></worksheet>");
}
