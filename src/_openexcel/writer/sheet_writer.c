#include "sheet_writer.h"
#include "../cell.h"
#include <stdint.h>
#include <string.h>
#include <stdio.h>
#include <ctype.h>
#include <stdlib.h>

static void emit_cell_ref(OxlXmlBuf *b, uint32_t row, uint16_t col) {
    oxl_xmlbuf_col_label(b, col);
    oxl_xmlbuf_uint(b, row + 1);  /* xlsx is 1-based */
}

void oxl_write_sheet(OxlXmlBuf *b, const OxlWorksheet *ws, const OxlWorkbook *wb,
                     uint16_t date_xf_idx) {
    (void)wb;

    /* ── XML declaration + worksheet opening tag ───────────────────────── */
    oxl_xmlbuf_cstr(b,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

    /* ── Feature D: <sheetPr> / tab color ─────────────────────────────── */
    if (ws->tab_color[0] != '\0') {
        oxl_xmlbuf_cstr(b, "<sheetPr><tabColor rgb=\"");
        oxl_xmlbuf_cstr(b, ws->tab_color);
        oxl_xmlbuf_cstr(b, "\"/></sheetPr>");
    }

    /* ── Feature C: <sheetViews> — always emit (Excel requires it) ─────── */
    oxl_xmlbuf_cstr(b, "<sheetViews><sheetView");
    if (!ws->show_gridlines) {
        oxl_xmlbuf_cstr(b, " showGridLines=\"0\"");
    }
    if (ws->zoom_scale > 0 && ws->zoom_scale != 100) {
        oxl_xmlbuf_cstr(b, " zoomScale=\"");
        oxl_xmlbuf_uint(b, (uint32_t)ws->zoom_scale);
        oxl_xmlbuf_raw(b, "\"", 1);
    }
    oxl_xmlbuf_cstr(b, " workbookViewId=\"0\">");

    if (ws->freeze_panes) {
        /* Parse "B2" into 0-based col/row offsets */
        const char *fp = ws->freeze_panes;
        uint32_t fc = 0;
        while (*fp && isalpha((unsigned char)*fp)) {
            fc = fc * 26 + (uint32_t)(toupper((unsigned char)*fp) - 'A' + 1);
            fp++;
        }
        uint16_t fp_col = (uint16_t)(fc > 0 ? fc - 1 : 0);  /* 0-based */
        uint32_t fp_row = (uint32_t)(atoi(fp) > 0 ? atoi(fp) - 1 : 0);  /* 0-based */

        oxl_xmlbuf_cstr(b, "<pane");
        if (fp_col > 0) {
            oxl_xmlbuf_cstr(b, " xSplit=\"");
            oxl_xmlbuf_uint(b, fp_col);
            oxl_xmlbuf_raw(b, "\"", 1);
        }
        if (fp_row > 0) {
            oxl_xmlbuf_cstr(b, " ySplit=\"");
            oxl_xmlbuf_uint(b, fp_row);
            oxl_xmlbuf_raw(b, "\"", 1);
        }
        oxl_xmlbuf_cstr(b, " topLeftCell=\"");
        oxl_xmlbuf_cstr(b, ws->freeze_panes);
        oxl_xmlbuf_cstr(b, "\" activePane=\"bottomLeft\" state=\"frozen\"/>");
    }
    oxl_xmlbuf_cstr(b, "</sheetView></sheetViews>");

    /* ── Feature A: <cols> block ─────────────────────────────────────────── */
    if (ws->col_dim_count > 0) {
        oxl_xmlbuf_cstr(b, "<cols>");
        for (uint32_t i = 0; i < ws->col_dim_count; i++) {
            const OxlColDim *d = &ws->col_dims[i];
            oxl_xmlbuf_cstr(b, "<col min=\"");
            oxl_xmlbuf_uint(b, d->col_min);
            oxl_xmlbuf_cstr(b, "\" max=\"");
            oxl_xmlbuf_uint(b, d->col_max);
            if (d->width > 0.0) {
                oxl_xmlbuf_cstr(b, "\" width=\"");
                oxl_xmlbuf_double(b, d->width);
            }
            if (d->custom_width) oxl_xmlbuf_cstr(b, "\" customWidth=\"1");
            if (d->best_fit)     oxl_xmlbuf_cstr(b, "\" bestFit=\"1");
            if (d->hidden)       oxl_xmlbuf_cstr(b, "\" hidden=\"1");
            oxl_xmlbuf_cstr(b, "\"/>");
        }
        oxl_xmlbuf_cstr(b, "</cols>");
    }

    /* ── <sheetData> ─────────────────────────────────────────────────────── */
    oxl_xmlbuf_cstr(b, "<sheetData>");

    uint32_t i = 0;
    while (i < ws->cell_count) {
        uint32_t row = ws->cells[i].row;

        /* Open row — with optional Feature A row-dim attributes */
        oxl_xmlbuf_cstr(b, "<row r=\"");
        oxl_xmlbuf_uint(b, row + 1);
        oxl_xmlbuf_raw(b, "\"", 1);

        /* Feature A: look up row dimension (linear scan; row dims are small) */
        uint32_t r1 = row + 1; /* 1-based */
        for (uint32_t d = 0; d < ws->row_dim_count; d++) {
            if (ws->row_dims[d].row_idx == r1) {
                if (ws->row_dims[d].height > 0.0) {
                    oxl_xmlbuf_cstr(b, " ht=\"");
                    oxl_xmlbuf_double(b, ws->row_dims[d].height);
                    oxl_xmlbuf_cstr(b, "\" customHeight=\"1\"");
                }
                if (ws->row_dims[d].hidden) {
                    oxl_xmlbuf_cstr(b, " hidden=\"1\"");
                }
                break;
            }
        }

        oxl_xmlbuf_raw(b, ">", 1);

        /* Emit all cells in this row */
        while (i < ws->cell_count && ws->cells[i].row == row) {
            const OxlCell *c = &ws->cells[i];
            i++;

            if (c->type == OXL_CELL_EMPTY && !c->formula && c->style_idx == 0) continue;

            /* Determine which style index to emit */
            uint16_t emit_style = c->style_idx;
            if (c->type == OXL_CELL_DATE && emit_style == 0)
                emit_style = date_xf_idx;  /* fallback to default date xf */

            oxl_xmlbuf_cstr(b, "<c r=\"");
            emit_cell_ref(b, c->row, c->col);
            oxl_xmlbuf_raw(b, "\"", 1);
            if (emit_style > 0) {
                oxl_xmlbuf_cstr(b, " s=\"");
                oxl_xmlbuf_uint(b, emit_style);
                oxl_xmlbuf_raw(b, "\"", 1);
            }

            /* Type attribute (t=) */
            switch (c->type) {
            case OXL_CELL_STRING:
            case OXL_CELL_INLINE_STR:
                oxl_xmlbuf_cstr(b, " t=\"s\"");
                break;
            case OXL_CELL_BOOL:
                oxl_xmlbuf_cstr(b, " t=\"b\"");
                break;
            case OXL_CELL_ERROR:
                oxl_xmlbuf_cstr(b, " t=\"e\"");
                break;
            default:
                break;
            }

            /* Close opening tag */
            oxl_xmlbuf_raw(b, ">", 1);

            /* Formula element (must come before <v> per OOXML spec) */
            if (c->formula) {
                oxl_xmlbuf_cstr(b, "<f>");
                oxl_xmlbuf_text(b, c->formula);
                oxl_xmlbuf_cstr(b, "</f>");
            }

            /* Value element (type-specific content) */
            switch (c->type) {
            case OXL_CELL_STRING:
            case OXL_CELL_INLINE_STR:
                oxl_xmlbuf_cstr(b, "<v>");
                oxl_xmlbuf_uint(b, c->v.s_idx);
                oxl_xmlbuf_cstr(b, "</v>");
                break;
            case OXL_CELL_BOOL:
                oxl_xmlbuf_cstr(b, "<v>");
                oxl_xmlbuf_raw(b, c->v.b ? "1" : "0", 1);
                oxl_xmlbuf_cstr(b, "</v>");
                break;
            case OXL_CELL_FLOAT:
                oxl_xmlbuf_cstr(b, "<v>");
                oxl_xmlbuf_double(b, c->v.f);
                oxl_xmlbuf_cstr(b, "</v>");
                break;
            case OXL_CELL_DATE: {
                double serial = oxl_date_to_serial(c->v.dt, wb->date1904);
                oxl_xmlbuf_cstr(b, "<v>");
                oxl_xmlbuf_double(b, serial);
                oxl_xmlbuf_cstr(b, "</v>");
                break;
            }
            case OXL_CELL_ERROR:
                oxl_xmlbuf_cstr(b, "<v>");
                oxl_xmlbuf_text(b, c->v.s_inline ? c->v.s_inline : "#ERR!");
                oxl_xmlbuf_cstr(b, "</v>");
                break;
            default:
                /* OXL_CELL_EMPTY with formula: no <v> element */
                break;
            }

            oxl_xmlbuf_cstr(b, "</c>");
        }
        oxl_xmlbuf_cstr(b, "</row>");
    }
    oxl_xmlbuf_cstr(b, "</sheetData>");

    /* ── Feature B: <mergeCells> ─────────────────────────────────────────── */
    if (ws->merged_cell_count > 0) {
        oxl_xmlbuf_cstr(b, "<mergeCells count=\"");
        oxl_xmlbuf_uint(b, ws->merged_cell_count);
        oxl_xmlbuf_cstr(b, "\">");
        for (uint32_t j = 0; j < ws->merged_cell_count; j++) {
            const OxlMergedCell *m = &ws->merged_cells[j];
            oxl_xmlbuf_cstr(b, "<mergeCell ref=\"");
            oxl_xmlbuf_col_label(b, m->min_col - 1);  /* 0-based for col_label */
            oxl_xmlbuf_uint(b, m->min_row);
            oxl_xmlbuf_raw(b, ":", 1);
            oxl_xmlbuf_col_label(b, m->max_col - 1);
            oxl_xmlbuf_uint(b, m->max_row);
            oxl_xmlbuf_cstr(b, "\"/>");
        }
        oxl_xmlbuf_cstr(b, "</mergeCells>");
    }

    /* ── Feature E: <autoFilter> ─────────────────────────────────────────── */
    if (ws->auto_filter_ref) {
        oxl_xmlbuf_cstr(b, "<autoFilter ref=\"");
        oxl_xmlbuf_text(b, ws->auto_filter_ref);
        oxl_xmlbuf_cstr(b, "\"/>");
    }

    /* Phase 8: <hyperlinks> */
    {
        int has_hyperlinks = 0;
        for (uint32_t hi = 0; hi < ws->cell_count; hi++) {
            if (ws->cells[hi].hyperlink) { has_hyperlinks = 1; break; }
        }
        if (has_hyperlinks) {
            oxl_xmlbuf_cstr(b, "<hyperlinks>");
            uint32_t rid = 1;
            for (uint32_t hi = 0; hi < ws->cell_count; hi++) {
                if (!ws->cells[hi].hyperlink) continue;
                oxl_xmlbuf_cstr(b, "<hyperlink ref=\"");
                emit_cell_ref(b, ws->cells[hi].row, ws->cells[hi].col);
                oxl_xmlbuf_cstr(b, "\" r:id=\"rId");
                oxl_xmlbuf_uint(b, rid++);
                oxl_xmlbuf_cstr(b, "\"/>");
            }
            oxl_xmlbuf_cstr(b, "</hyperlinks>");
        }
    }

    oxl_xmlbuf_cstr(b, "</worksheet>");
}
