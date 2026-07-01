#include "xml_sheet.h"
#include "xml_sheet_rels.h"
#include "../cell.h"
#include "../styles.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <ctype.h>

typedef enum {
    SS_NONE,
    SS_SHEET_DATA,
    SS_ROW,
    SS_CELL,
    SS_VALUE,
    SS_FORMULA,
    SS_INLINE_SI,
    SS_INLINE_T,
    /* Feature A */
    SS_COLS,
    /* Feature C */
    SS_SHEET_VIEWS,
    SS_SHEET_VIEW,
    /* Phase 8: hyperlinks */
    SS_HYPERLINKS,
    /* Phase 13: data validations */
    SS_DATA_VALIDATIONS,
    SS_DV_FORMULA,       /* inside <formula1> or <formula2>, collects text */
    /* Phase 14: page setup */
    SS_PRINT_OPTIONS,      /* inside <printOptions> — attributes only, no children */
    SS_PAGE_MARGINS,       /* inside <pageMargins> — attributes only */
    SS_PAGE_SETUP,         /* inside <pageSetup> — attributes only */
    /* Phase 15: sheet protection */
    SS_SHEET_PROTECTION,   /* inside <sheetProtection> — attributes only */
    /* Phase 16: conditional formatting */
    SS_COND_FMT,           /* inside <conditionalFormatting> */
    SS_CF_RULE,            /* inside <cfRule> */
    SS_CF_FORMULA,         /* inside <formula> inside <cfRule> */
    SS_CF_COLOR_SCALE,     /* inside <colorScale> */
    SS_CF_DATA_BAR,        /* inside <dataBar> */
} SheetState;

#define CBUF_STACK 512

typedef struct {
    OxlWorksheet *ws;
    OxlWorkbook  *wb;
    SheetState    state;
    /* current cell accumulator */
    uint32_t      cur_row;
    uint16_t      cur_col;
    uint16_t      cur_style;
    char          cur_type;  /* 's', 'b', 'e', 'n', 0=numeric */
    int           cur_is_inline;
    /* character data buffer */
    char          stack_buf[CBUF_STACK];
    char         *heap_buf;
    char         *cbuf;     /* points to stack_buf or heap_buf */
    size_t        cbuf_len;
    size_t        cbuf_cap;
    /* formula buffer (separate from cbuf so <v> can follow <f>) */
    char          formula_stack[CBUF_STACK];
    char         *formula_heap;
    char         *formula_buf;  /* points to formula_stack or formula_heap */
    size_t        formula_len;
    size_t        formula_cap;
    /* rich inline string accumulator */
    char         *rich_buf;
    size_t        rich_len;
    size_t        rich_cap;
    int           error;
    /* Phase 8: hyperlink rels (NULL if no rels file for this sheet) */
    const OxlHyperlinkRels *rels;
    /* Phase 13: data validation accumulator */
    OxlDataValidation cur_dv;
    int               cur_dv_formula_slot; /* 1 = formula1, 2 = formula2 */
    /* Phase 16: conditional formatting accumulator */
    OxlCfRule    cur_cf_rule;
    char         cur_cf_sqref[256];
    int          cur_cf_formula_slot; /* 1 = formula, 2 = formula2 */
    int          cur_cf_in_color_scale;
    int          cur_cf_in_data_bar;
} SheetCtx;

/* Parse "B5" → row=4, col=1 */
static void parse_cell_ref(const char *r, uint32_t *row, uint16_t *col) {
    uint32_t c = 0;
    while (*r && isalpha((unsigned char)*r)) {
        c = c * 26 + (uint32_t)(toupper((unsigned char)*r) - 'A' + 1);
        r++;
    }
    *col = (uint16_t)(c > 0 ? c - 1 : 0);
    *row = (uint32_t)(atoi(r) > 0 ? atoi(r) - 1 : 0);
}

static const char *attr(const char **attrs, const char *key) {
    for (; *attrs; attrs += 2) if (strcmp(attrs[0], key) == 0) return attrs[1];
    return NULL;
}

static void cbuf_reset(SheetCtx *c) {
    c->cbuf     = c->stack_buf;
    c->cbuf_len = 0;
    c->cbuf_cap = CBUF_STACK;
}

static void cbuf_append(SheetCtx *c, const char *s, int n) {
    if ((size_t)n + c->cbuf_len >= c->cbuf_cap) {
        size_t newcap = c->cbuf_cap * 2 + (size_t)n;
        char *p = realloc(c->cbuf == c->stack_buf ? NULL : c->heap_buf, newcap);
        if (!p) { c->error = 1; return; }
        if (c->cbuf == c->stack_buf) memcpy(p, c->stack_buf, c->cbuf_len);
        c->heap_buf = p;
        c->cbuf     = p;
        c->cbuf_cap = newcap;
    }
    memcpy(c->cbuf + c->cbuf_len, s, (size_t)n);
    c->cbuf_len += (size_t)n;
}

static void formula_buf_reset(SheetCtx *c) {
    c->formula_buf = c->formula_stack;
    c->formula_len = 0;
    c->formula_cap = CBUF_STACK;
}

static void formula_buf_append(SheetCtx *c, const char *s, int n) {
    if ((size_t)n + c->formula_len >= c->formula_cap) {
        size_t newcap = c->formula_cap * 2 + (size_t)n;
        char *p = realloc(c->formula_buf == c->formula_stack ? NULL : c->formula_heap, newcap);
        if (!p) { c->error = 1; return; }
        if (c->formula_buf == c->formula_stack) memcpy(p, c->formula_stack, c->formula_len);
        c->formula_heap = p;
        c->formula_buf  = p;
        c->formula_cap  = newcap;
    }
    memcpy(c->formula_buf + c->formula_len, s, (size_t)n);
    c->formula_len += (size_t)n;
}

static void rich_append(SheetCtx *c, const char *s, size_t n) {
    if (c->rich_len + n >= c->rich_cap) {
        size_t newcap = c->rich_cap ? c->rich_cap * 2 + n : 256;
        char *p = realloc(c->rich_buf, newcap);
        if (!p) return;
        c->rich_buf = p;
        c->rich_cap = newcap;
    }
    memcpy(c->rich_buf + c->rich_len, s, n);
    c->rich_len += n;
}

static void emit_cell(SheetCtx *c) {
    if (c->cbuf_len == 0 && !c->cur_is_inline && c->formula_len == 0 && c->cur_style == 0) return;

    OxlCell cell;
    memset(&cell, 0, sizeof(cell));
    cell.row       = c->cur_row;
    cell.col       = c->cur_col;
    cell.style_idx = c->cur_style;

    /* capture formula if present (raw expression without '=' prefix) */
    if (c->formula_len > 0) {
        if (c->formula_len < c->formula_cap) c->formula_buf[c->formula_len] = '\0';
        cell.formula = strndup(c->formula_buf, c->formula_len);
    }

    /* null-terminate cbuf */
    if (c->cbuf_len < c->cbuf_cap) c->cbuf[c->cbuf_len] = '\0';

    if (c->cbuf_len == 0 && !c->cur_is_inline) {
        /* Formula with no cached <v> value — store as EMPTY with formula */
        cell.type = OXL_CELL_EMPTY;
    } else if (c->cur_is_inline) {
        /* inline string: copy rich_buf or cbuf */
        const char *text = c->rich_len ? c->rich_buf : c->cbuf;
        size_t tlen = c->rich_len ? c->rich_len : c->cbuf_len;
        cell.type = OXL_CELL_INLINE_STR;
        cell.v.s_inline = malloc(tlen + 1);
        if (cell.v.s_inline) {
            memcpy(cell.v.s_inline, text, tlen);
            cell.v.s_inline[tlen] = '\0';
        }
    } else if (c->cur_type == 's') {
        cell.type  = OXL_CELL_STRING;
        cell.v.s_idx = (uint32_t)atoi(c->cbuf);
    } else if (c->cur_type == 'b') {
        cell.type  = OXL_CELL_BOOL;
        cell.v.b   = (c->cbuf[0] == '1') ? 1 : 0;
    } else if (c->cur_type == 'e') {
        cell.type  = OXL_CELL_ERROR;
        cell.v.s_inline = strndup(c->cbuf, c->cbuf_len);
    } else {
        /* numeric (default) */
        double val = atof(c->cbuf);
        if (oxl_styles_is_date(&c->wb->styles, c->cur_style)) {
            cell.type = OXL_CELL_DATE;
            cell.v.dt = oxl_serial_to_date(val, c->wb->date1904);
        } else {
            cell.type = OXL_CELL_FLOAT;
            cell.v.f  = val;
        }
    }

    oxl_worksheet_append_cell(c->ws, &cell);
    c->cur_is_inline = 0;
    c->rich_len = 0;
}

static void XMLCALL sheet_start(void *ud, const char *name, const char **attrs) {
    SheetCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    switch (c->state) {
    case SS_NONE:
        /* Feature C: sheetViews */
        if (strcmp(name, "sheetViews") == 0) {
            c->state = SS_SHEET_VIEWS;
        }
        /* Feature A: cols block */
        else if (strcmp(name, "cols") == 0) {
            c->state = SS_COLS;
        }
        /* Feature D: tab color (sheetPr/tabColor) */
        else if (strcmp(name, "tabColor") == 0) {
            const char *rgb = attr(attrs, "rgb");
            const char *theme = attr(attrs, "theme");
            if (rgb) {
                strncpy(c->ws->tab_color, rgb, 8);
                c->ws->tab_color[8] = '\0';
            }
            (void)theme;
        }
        /* Feature B: mergeCell */
        else if (strcmp(name, "mergeCell") == 0) {
            const char *ref = attr(attrs, "ref");
            if (ref) {
                char buf[32];
                strncpy(buf, ref, 31);
                buf[31] = '\0';
                char *colon = strchr(buf, ':');
                if (colon) {
                    *colon = '\0';
                    uint32_t r1, r2;
                    uint16_t c1, c2;
                    parse_cell_ref(buf, &r1, &c1);
                    parse_cell_ref(colon + 1, &r2, &c2);
                    oxl_worksheet_add_merge(c->ws, r1 + 1, c1 + 1, r2 + 1, c2 + 1);
                }
            }
        }
        /* Feature E: autoFilter */
        else if (strcmp(name, "autoFilter") == 0) {
            const char *ref = attr(attrs, "ref");
            if (ref) {
                free(c->ws->auto_filter_ref);
                c->ws->auto_filter_ref = strdup(ref);
            }
        }
        else if (strcmp(name, "sheetData") == 0) {
            c->state = SS_SHEET_DATA;
        }
        /* Phase 8: hyperlinks block */
        else if (strcmp(name, "hyperlinks") == 0) {
            c->state = SS_HYPERLINKS;
        }
        /* Phase 13: data validations block */
        else if (strcmp(name, "dataValidations") == 0) {
            c->state = SS_DATA_VALIDATIONS;
        }
        /* Phase 14: print options, page margins, page setup */
        else if (strcmp(name, "printOptions") == 0) {
            const char *gl  = attr(attrs, "gridLines");
            const char *hd  = attr(attrs, "headings");
            const char *hc  = attr(attrs, "horizontalCentered");
            const char *vc  = attr(attrs, "verticalCentered");
            if (gl && (gl[0] == '1' || strcmp(gl, "true") == 0))
                c->ws->print_options.grid_lines = 1;
            if (hd && (hd[0] == '1' || strcmp(hd, "true") == 0))
                c->ws->print_options.headings = 1;
            if (hc && (hc[0] == '1' || strcmp(hc, "true") == 0))
                c->ws->print_options.horizontal_centered = 1;
            if (vc && (vc[0] == '1' || strcmp(vc, "true") == 0))
                c->ws->print_options.vertical_centered = 1;
            if (gl || hd || hc || vc)
                c->ws->print_options.has_options = 1;
            c->state = SS_PRINT_OPTIONS;
        }
        else if (strcmp(name, "pageMargins") == 0) {
            const char *l = attr(attrs, "left");
            const char *r = attr(attrs, "right");
            const char *t = attr(attrs, "top");
            const char *b = attr(attrs, "bottom");
            const char *h = attr(attrs, "header");
            const char *f = attr(attrs, "footer");
            if (l) c->ws->page_margins.left   = atof(l);
            if (r) c->ws->page_margins.right  = atof(r);
            if (t) c->ws->page_margins.top    = atof(t);
            if (b) c->ws->page_margins.bottom = atof(b);
            if (h) c->ws->page_margins.header = atof(h);
            if (f) c->ws->page_margins.footer = atof(f);
            if (l || r || t || b || h || f)
                c->ws->page_margins.has_margins = 1;
            c->state = SS_PAGE_MARGINS;
        }
        else if (strcmp(name, "pageSetup") == 0) {
            const char *orient = attr(attrs, "orientation");
            const char *ps     = attr(attrs, "paperSize");
            const char *sc     = attr(attrs, "scale");
            const char *fw     = attr(attrs, "fitToWidth");
            const char *fh     = attr(attrs, "fitToHeight");
            const char *fp     = attr(attrs, "fitToPage");
            if (orient) {
                free(c->ws->page_setup.orientation);
                c->ws->page_setup.orientation = strdup(orient);
            }
            if (ps) c->ws->page_setup.paper_size    = (uint32_t)atoi(ps);
            if (sc) c->ws->page_setup.scale         = (uint32_t)atoi(sc);
            if (fw) c->ws->page_setup.fit_to_width  = (uint32_t)atoi(fw);
            if (fh) c->ws->page_setup.fit_to_height = (uint32_t)atoi(fh);
            if (fp && (fp[0] == '1' || strcmp(fp, "true") == 0))
                c->ws->page_setup.fit_to_page = 1;
            if (orient || ps || sc || fw || fh || fp)
                c->ws->page_setup.has_setup = 1;
            c->state = SS_PAGE_SETUP;
        }
        /* Phase 16: conditional formatting */
        else if (strcmp(name, "conditionalFormatting") == 0) {
            const char *sqref_s = attr(attrs, "sqref");
            if (sqref_s) {
                strncpy(c->cur_cf_sqref, sqref_s, sizeof(c->cur_cf_sqref) - 1);
                c->cur_cf_sqref[sizeof(c->cur_cf_sqref) - 1] = '\0';
            } else {
                c->cur_cf_sqref[0] = '\0';
            }
            c->state = SS_COND_FMT;
        }
        /* Phase 15: sheet protection */
        else if (strcmp(name, "sheetProtection") == 0) {
            OxlSheetProtection *p = &c->ws->protection;
            #define BOOL_ATTR(s) ((s) && ((s)[0] == '1' || strcmp((s), "true") == 0))
            const char *sheet_s = attr(attrs, "sheet");
            const char *obj_s   = attr(attrs, "objects");
            const char *scen_s  = attr(attrs, "scenarios");
            const char *fc_s    = attr(attrs, "formatCells");
            const char *fcol_s  = attr(attrs, "formatColumns");
            const char *fr_s    = attr(attrs, "formatRows");
            const char *ic_s    = attr(attrs, "insertColumns");
            const char *ir_s    = attr(attrs, "insertRows");
            const char *ih_s    = attr(attrs, "insertHyperlinks");
            const char *dc_s    = attr(attrs, "deleteColumns");
            const char *dr_s    = attr(attrs, "deleteRows");
            const char *sl_s    = attr(attrs, "selectLockedCells");
            const char *sort_s  = attr(attrs, "sort");
            const char *af_s    = attr(attrs, "autoFilter");
            const char *pt_s    = attr(attrs, "pivotTables");
            const char *su_s    = attr(attrs, "selectUnlockedCells");
            const char *algo_s  = attr(attrs, "algorithmName");
            const char *hash_s  = attr(attrs, "hashValue");
            const char *salt_s  = attr(attrs, "saltValue");
            const char *spin_s  = attr(attrs, "spinCount");
            const char *pwd_s   = attr(attrs, "password");
            p->sheet               = BOOL_ATTR(sheet_s);
            p->objects             = BOOL_ATTR(obj_s);
            p->scenarios           = BOOL_ATTR(scen_s);
            p->format_cells        = !BOOL_ATTR(fc_s);
            p->format_columns      = !BOOL_ATTR(fcol_s);
            p->format_rows         = !BOOL_ATTR(fr_s);
            p->insert_columns      = !BOOL_ATTR(ic_s);
            p->insert_rows         = !BOOL_ATTR(ir_s);
            p->insert_hyperlinks   = !BOOL_ATTR(ih_s);
            p->delete_columns      = !BOOL_ATTR(dc_s);
            p->delete_rows         = !BOOL_ATTR(dr_s);
            p->select_locked       = !BOOL_ATTR(sl_s);
            p->sort                = !BOOL_ATTR(sort_s);
            p->auto_filter         = !BOOL_ATTR(af_s);
            p->pivot_tables        = !BOOL_ATTR(pt_s);
            p->select_unlocked     = BOOL_ATTR(su_s);
            #undef BOOL_ATTR
            if (algo_s) { free(p->algorithm_name); p->algorithm_name = strdup(algo_s); }
            if (hash_s) { free(p->hash_value);     p->hash_value     = strdup(hash_s); }
            if (salt_s) { free(p->salt_value);     p->salt_value     = strdup(salt_s); }
            if (spin_s) p->spin_count = (uint32_t)atoi(spin_s);
            if (pwd_s)  { free(p->password_hash);  p->password_hash  = strdup(pwd_s); }
            p->has_protection = 1;
            c->state = SS_SHEET_PROTECTION;
        }
        break;

    /* Phase 8: hyperlink elements */
    case SS_HYPERLINKS:
        if (strcmp(name, "hyperlink") == 0) {
            const char *ref = attr(attrs, "ref");
            /* r:id attribute — expat passes namespace prefix literally */
            const char *rid = attr(attrs, "r:id");
            if (!rid) rid = attr(attrs, "id");  /* fallback */
            if (ref && rid && c->rels) {
                const char *url = oxl_hyperlink_rels_lookup(c->rels, rid);
                if (url) {
                    uint32_t row; uint16_t col;
                    parse_cell_ref(ref, &row, &col);
                    /* Binary search for existing cell */
                    uint64_t key = ((uint64_t)row << 16) | col;
                    uint32_t lo = 0, hi = c->ws->cell_count;
                    while (lo < hi) {
                        uint32_t mid = lo + (hi - lo) / 2;
                        uint64_t mk = ((uint64_t)c->ws->cells[mid].row << 16) | c->ws->cells[mid].col;
                        if (mk == key) {
                            free(c->ws->cells[mid].hyperlink);
                            c->ws->cells[mid].hyperlink = strdup(url);
                            break;
                        }
                        if (mk < key) lo = mid + 1; else hi = mid;
                    }
                }
            }
        }
        break;

    /* Phase 13: data validation elements */
    case SS_DATA_VALIDATIONS:
        if (strcmp(name, "dataValidation") == 0) {
            memset(&c->cur_dv, 0, sizeof(c->cur_dv));
            c->cur_dv_formula_slot = 0;
            const char *type_s    = attr(attrs, "type");
            const char *op_s      = attr(attrs, "operator");
            const char *sqref_s   = attr(attrs, "sqref");
            const char *ab_s      = attr(attrs, "allowBlank");
            const char *sdd_s     = attr(attrs, "showDropDown");
            const char *sem_s     = attr(attrs, "showErrorMessage");
            const char *sim_s     = attr(attrs, "showInputMessage");
            const char *et_s      = attr(attrs, "errorTitle");
            const char *em_s      = attr(attrs, "error");
            const char *es_s      = attr(attrs, "errorStyle");
            const char *pt_s      = attr(attrs, "promptTitle");
            const char *pm_s      = attr(attrs, "prompt");
            if (type_s)  c->cur_dv.type          = strdup(type_s);
            if (op_s)    c->cur_dv.dv_operator    = strdup(op_s);
            if (sqref_s) c->cur_dv.sqref          = strdup(sqref_s);
            c->cur_dv.allow_blank        = ab_s  && (ab_s[0]  == '1' || strcmp(ab_s,  "true") == 0);
            c->cur_dv.show_drop_down     = sdd_s && (sdd_s[0] == '1' || strcmp(sdd_s, "true") == 0);
            c->cur_dv.show_error_message = sem_s && (sem_s[0] == '1' || strcmp(sem_s, "true") == 0);
            c->cur_dv.show_input_message = sim_s && (sim_s[0] == '1' || strcmp(sim_s, "true") == 0);
            if (et_s) c->cur_dv.error_title    = strdup(et_s);
            if (em_s) c->cur_dv.error_message  = strdup(em_s);
            if (es_s) c->cur_dv.error_style    = strdup(es_s);
            if (pt_s) c->cur_dv.prompt_title   = strdup(pt_s);
            if (pm_s) c->cur_dv.prompt_message = strdup(pm_s);
        } else if (strcmp(name, "formula1") == 0) {
            cbuf_reset(c);
            c->cur_dv_formula_slot = 1;
            c->state = SS_DV_FORMULA;
        } else if (strcmp(name, "formula2") == 0) {
            cbuf_reset(c);
            c->cur_dv_formula_slot = 2;
            c->state = SS_DV_FORMULA;
        }
        break;

    case SS_DV_FORMULA:
        /* no children expected */
        break;

    /* Feature C: sheetViews / sheetView / pane */
    case SS_SHEET_VIEWS:
        if (strcmp(name, "sheetView") == 0) {
            const char *zs = attr(attrs, "zoomScale");
            const char *gl = attr(attrs, "showGridLines");
            if (zs) c->ws->zoom_scale = atoi(zs);
            if (gl && (gl[0] == '0' || strcmp(gl, "false") == 0))
                c->ws->show_gridlines = 0;
            c->state = SS_SHEET_VIEW;
        }
        break;

    case SS_SHEET_VIEW:
        if (strcmp(name, "pane") == 0) {
            const char *state_s  = attr(attrs, "state");
            const char *top_left = attr(attrs, "topLeftCell");
            if (state_s && strcmp(state_s, "frozen") == 0 && top_left) {
                free(c->ws->freeze_panes);
                c->ws->freeze_panes = strdup(top_left);
            }
        }
        break;

    /* Feature A: col element */
    case SS_COLS:
        if (strcmp(name, "col") == 0) {
            const char *min_s = attr(attrs, "min");
            const char *max_s = attr(attrs, "max");
            const char *w_s   = attr(attrs, "width");
            const char *hid_s = attr(attrs, "hidden");
            const char *bf_s  = attr(attrs, "bestFit");
            const char *cw_s  = attr(attrs, "customWidth");
            uint16_t mn = min_s ? (uint16_t)atoi(min_s) : 1;
            uint16_t mx = max_s ? (uint16_t)atoi(max_s) : mn;
            double w    = w_s   ? atof(w_s)  : 0.0;
            int hid = hid_s && (hid_s[0] == '1' || strcmp(hid_s, "true") == 0);
            int bf  = bf_s  && (bf_s[0]  == '1' || strcmp(bf_s,  "true") == 0);
            int cw  = cw_s  && (cw_s[0]  == '1' || strcmp(cw_s,  "true") == 0);
            oxl_worksheet_set_col_dim(c->ws, mn, mx, w, hid, bf, cw);
        }
        break;

    case SS_SHEET_DATA:
        if (strcmp(name, "row") == 0) {
            const char *r = attr(attrs, "r");
            c->cur_row = r ? (uint32_t)(atoi(r) - 1) : c->cur_row;

            /* Feature A: row height / hidden */
            const char *ht_s  = attr(attrs, "ht");
            const char *hid_s = attr(attrs, "hidden");
            const char *ch_s  = attr(attrs, "customHeight");
            if (ht_s || hid_s) {
                double h = ht_s ? atof(ht_s) : 0.0;
                int hid  = hid_s && (hid_s[0] == '1' || strcmp(hid_s, "true") == 0);
                int ch   = ch_s  && (ch_s[0]  == '1' || strcmp(ch_s,  "true") == 0);
                uint32_t r1 = c->cur_row + 1; /* convert 0-based to 1-based */
                oxl_worksheet_set_row_dim(c->ws, r1, h, hid, ch);
            }

            c->state = SS_ROW;
        }
        break;

    case SS_ROW:
        if (strcmp(name, "c") == 0) {
            const char *r = attr(attrs, "r");
            const char *s = attr(attrs, "s");
            const char *t = attr(attrs, "t");
            if (r) parse_cell_ref(r, &c->cur_row, &c->cur_col);
            c->cur_style = s ? (uint16_t)atoi(s) : 0;
            c->cur_type  = t ? t[0] : 0;
            c->cur_is_inline = (t && strcmp(t, "inlineStr") == 0) ? 1 : 0;
            cbuf_reset(c);
            c->state = SS_CELL;
        }
        break;

    case SS_CELL:
        if (strcmp(name, "v") == 0) {
            cbuf_reset(c);
            c->state = SS_VALUE;
        } else if (strcmp(name, "f") == 0) {
            c->state = SS_FORMULA;
        } else if (strcmp(name, "is") == 0) {
            c->rich_len = 0;
            c->state = SS_INLINE_SI;
        }
        break;

    case SS_INLINE_SI:
        if (strcmp(name, "t") == 0) {
            cbuf_reset(c);
            c->state = SS_INLINE_T;
        }
        break;

    /* Phase 16: conditional formatting */
    case SS_COND_FMT:
        if (strcmp(name, "cfRule") == 0) {
            memset(&c->cur_cf_rule, 0, sizeof(c->cur_cf_rule));
            c->cur_cf_rule.dxf_id = -1;
            c->cur_cf_in_color_scale = 0;
            c->cur_cf_in_data_bar = 0;
            const char *type_s  = attr(attrs, "type");
            const char *op_s    = attr(attrs, "operator");
            const char *pri_s   = attr(attrs, "priority");
            const char *dxf_s   = attr(attrs, "dxfId");
            const char *text_s  = attr(attrs, "text");
            const char *sit_s   = attr(attrs, "stopIfTrue");
            const char *rank_s  = attr(attrs, "rank");
            const char *top_s   = attr(attrs, "top");
            const char *pct_s   = attr(attrs, "percent");
            const char *aa_s    = attr(attrs, "aboveAverage");
            const char *ea_s    = attr(attrs, "equalAverage");
            if (type_s)  c->cur_cf_rule.type      = strdup(type_s);
            if (op_s)    c->cur_cf_rule.operator_  = strdup(op_s);
            if (text_s)  c->cur_cf_rule.text       = strdup(text_s);
            if (pri_s)   c->cur_cf_rule.priority   = atoi(pri_s);
            if (dxf_s)   c->cur_cf_rule.dxf_id    = atoi(dxf_s);
            if (sit_s && (strcmp(sit_s, "1") == 0 || strcmp(sit_s, "true") == 0))
                c->cur_cf_rule.stop_if_true = 1;
            if (rank_s)  c->cur_cf_rule.top10_rank    = (uint32_t)atoi(rank_s);
            if (top_s && (strcmp(top_s, "0") != 0))
                c->cur_cf_rule.top10_top = 1;
            else if (!top_s)
                c->cur_cf_rule.top10_top = 1; /* default: top */
            if (pct_s && (strcmp(pct_s, "1") == 0 || strcmp(pct_s, "true") == 0))
                c->cur_cf_rule.top10_percent = 1;
            /* aboveAverage: default=1 (above), explicit "0" means below */
            if (aa_s && strcmp(aa_s, "0") == 0)
                c->cur_cf_rule.above_avg = 0;
            else
                c->cur_cf_rule.above_avg = 1;
            if (ea_s && (strcmp(ea_s, "1") == 0 || strcmp(ea_s, "true") == 0))
                c->cur_cf_rule.equal_avg = 1;
            c->state = SS_CF_RULE;
        }
        break;

    case SS_CF_RULE:
        if (strcmp(name, "formula") == 0) {
            cbuf_reset(c);
            c->cur_cf_formula_slot = (c->cur_cf_rule.formula == NULL) ? 1 : 2;
            c->state = SS_CF_FORMULA;
        } else if (strcmp(name, "colorScale") == 0) {
            c->cur_cf_in_color_scale = 1;
            c->state = SS_CF_COLOR_SCALE;
        } else if (strcmp(name, "dataBar") == 0) {
            c->cur_cf_in_data_bar = 1;
            c->state = SS_CF_DATA_BAR;
        }
        break;

    case SS_CF_FORMULA:
        /* no children expected */
        break;

    case SS_CF_COLOR_SCALE:
    case SS_CF_DATA_BAR:
        if (strcmp(name, "cfvo") == 0) {
            if (c->cur_cf_rule.cfvo_count < 3) {
                uint32_t idx = c->cur_cf_rule.cfvo_count++;
                const char *type_s = attr(attrs, "type");
                const char *val_s  = attr(attrs, "val");
                c->cur_cf_rule.cfvos[idx].type    = type_s ? strdup(type_s) : NULL;
                c->cur_cf_rule.cfvos[idx].val     = val_s  ? strdup(val_s)  : NULL;
                c->cur_cf_rule.cfvos[idx].has_rgb = 0;
                c->cur_cf_rule.cfvos[idx].rgb     = 0;
            }
        } else if (strcmp(name, "color") == 0) {
            if (c->cur_cf_rule.color_count < 3) {
                const char *rgb_s = attr(attrs, "rgb");
                if (rgb_s) {
                    c->cur_cf_rule.colors[c->cur_cf_rule.color_count++] =
                        (uint32_t)strtoul(rgb_s, NULL, 16);
                }
            }
        }
        break;

    default:
        break;
    }
}

static void XMLCALL sheet_end(void *ud, const char *name) {
    SheetCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "v") == 0 && c->state == SS_VALUE) {
        c->state = SS_CELL;
    } else if (strcmp(name, "f") == 0 && c->state == SS_FORMULA) {
        /* formula text accumulated in formula_buf; keep it for emit_cell */
        c->state = SS_CELL;
    } else if (strcmp(name, "t") == 0 && c->state == SS_INLINE_T) {
        rich_append(c, c->cbuf, c->cbuf_len);
        c->state = SS_INLINE_SI;
    } else if (strcmp(name, "is") == 0 && c->state == SS_INLINE_SI) {
        c->state = SS_CELL;
    } else if (strcmp(name, "c") == 0 && c->state == SS_CELL) {
        emit_cell(c);
        formula_buf_reset(c);  /* clear formula for next cell */
        c->state = SS_ROW;
    } else if (strcmp(name, "row") == 0 && c->state == SS_ROW) {
        c->state = SS_SHEET_DATA;
    } else if (strcmp(name, "sheetData") == 0 && c->state == SS_SHEET_DATA) {
        c->state = SS_NONE;
    }
    /* Feature A: cols end */
    else if (strcmp(name, "cols") == 0 && c->state == SS_COLS) {
        c->state = SS_NONE;
    }
    /* Feature C: sheetView / sheetViews end */
    else if (strcmp(name, "sheetView") == 0 && c->state == SS_SHEET_VIEW) {
        c->state = SS_SHEET_VIEWS;
    } else if (strcmp(name, "sheetViews") == 0 && c->state == SS_SHEET_VIEWS) {
        c->state = SS_NONE;
    }
    /* Phase 8: hyperlinks block end */
    else if (strcmp(name, "hyperlinks") == 0 && c->state == SS_HYPERLINKS) {
        c->state = SS_NONE;
    }
    /* Phase 13: data validations end */
    else if (strcmp(name, "formula1") == 0 && c->state == SS_DV_FORMULA &&
             c->cur_dv_formula_slot == 1) {
        free(c->cur_dv.formula1);
        c->cur_dv.formula1 = c->cbuf_len > 0 ? strndup(c->cbuf, c->cbuf_len) : NULL;
        c->state = SS_DATA_VALIDATIONS;
    } else if (strcmp(name, "formula2") == 0 && c->state == SS_DV_FORMULA &&
               c->cur_dv_formula_slot == 2) {
        free(c->cur_dv.formula2);
        c->cur_dv.formula2 = c->cbuf_len > 0 ? strndup(c->cbuf, c->cbuf_len) : NULL;
        c->state = SS_DATA_VALIDATIONS;
    } else if (strcmp(name, "dataValidation") == 0 && c->state == SS_DATA_VALIDATIONS) {
        oxl_worksheet_add_data_validation(c->ws, &c->cur_dv);
        oxl_data_validation_free_fields(&c->cur_dv);
        memset(&c->cur_dv, 0, sizeof(c->cur_dv));
    } else if (strcmp(name, "dataValidations") == 0 && c->state == SS_DATA_VALIDATIONS) {
        c->state = SS_NONE;
    }
    /* Phase 14: page setup end — these are self-closing or no-child elements */
    else if (strcmp(name, "printOptions") == 0 && c->state == SS_PRINT_OPTIONS) {
        c->state = SS_NONE;
    }
    else if (strcmp(name, "pageMargins") == 0 && c->state == SS_PAGE_MARGINS) {
        c->state = SS_NONE;
    }
    else if (strcmp(name, "pageSetup") == 0 && c->state == SS_PAGE_SETUP) {
        c->state = SS_NONE;
    }
    /* Phase 15: sheet protection end */
    else if (strcmp(name, "sheetProtection") == 0 && c->state == SS_SHEET_PROTECTION) {
        c->state = SS_NONE;
    }
    /* Phase 16: conditional formatting end */
    else if (strcmp(name, "formula") == 0 && c->state == SS_CF_FORMULA) {
        if (c->cur_cf_formula_slot == 1) {
            free(c->cur_cf_rule.formula);
            c->cur_cf_rule.formula = c->cbuf_len > 0 ? strndup(c->cbuf, c->cbuf_len) : NULL;
        } else {
            free(c->cur_cf_rule.formula2);
            c->cur_cf_rule.formula2 = c->cbuf_len > 0 ? strndup(c->cbuf, c->cbuf_len) : NULL;
        }
        c->state = SS_CF_RULE;
    }
    else if (strcmp(name, "colorScale") == 0 && c->state == SS_CF_COLOR_SCALE) {
        c->cur_cf_in_color_scale = 0;
        c->state = SS_CF_RULE;
    }
    else if (strcmp(name, "dataBar") == 0 && c->state == SS_CF_DATA_BAR) {
        c->cur_cf_in_data_bar = 0;
        c->state = SS_CF_RULE;
    }
    else if (strcmp(name, "cfRule") == 0 && c->state == SS_CF_RULE) {
        /* Look up DXF if dxf_id >= 0 */
        if (c->cur_cf_rule.dxf_id >= 0) {
            const OxlDxf *dxf = oxl_styles_get_dxf(&c->wb->styles,
                                                     (uint32_t)c->cur_cf_rule.dxf_id);
            if (dxf) {
                if (dxf->font && !c->cur_cf_rule.font) {
                    c->cur_cf_rule.font = malloc(sizeof(OxlFontDef));
                    if (c->cur_cf_rule.font) {
                        *c->cur_cf_rule.font = *dxf->font;
                        c->cur_cf_rule.font->name = dxf->font->name ? strdup(dxf->font->name) : NULL;
                    }
                }
                if (dxf->fill && !c->cur_cf_rule.fill) {
                    c->cur_cf_rule.fill = malloc(sizeof(OxlFillDef));
                    if (c->cur_cf_rule.fill) {
                        *c->cur_cf_rule.fill = *dxf->fill;
                        c->cur_cf_rule.fill->pattern_type = dxf->fill->pattern_type ? strdup(dxf->fill->pattern_type) : NULL;
                    }
                }
                if (dxf->border && !c->cur_cf_rule.border) {
                    c->cur_cf_rule.border = malloc(sizeof(OxlBorderDef));
                    if (c->cur_cf_rule.border) {
                        c->cur_cf_rule.border->left.style     = dxf->border->left.style     ? strdup(dxf->border->left.style)     : NULL;
                        c->cur_cf_rule.border->left.color_rgb = dxf->border->left.color_rgb;
                        c->cur_cf_rule.border->left.has_color = dxf->border->left.has_color;
                        c->cur_cf_rule.border->right.style     = dxf->border->right.style     ? strdup(dxf->border->right.style)     : NULL;
                        c->cur_cf_rule.border->right.color_rgb = dxf->border->right.color_rgb;
                        c->cur_cf_rule.border->right.has_color = dxf->border->right.has_color;
                        c->cur_cf_rule.border->top.style     = dxf->border->top.style     ? strdup(dxf->border->top.style)     : NULL;
                        c->cur_cf_rule.border->top.color_rgb = dxf->border->top.color_rgb;
                        c->cur_cf_rule.border->top.has_color = dxf->border->top.has_color;
                        c->cur_cf_rule.border->bottom.style     = dxf->border->bottom.style     ? strdup(dxf->border->bottom.style)     : NULL;
                        c->cur_cf_rule.border->bottom.color_rgb = dxf->border->bottom.color_rgb;
                        c->cur_cf_rule.border->bottom.has_color = dxf->border->bottom.has_color;
                        c->cur_cf_rule.border->diagonal.style     = dxf->border->diagonal.style     ? strdup(dxf->border->diagonal.style)     : NULL;
                        c->cur_cf_rule.border->diagonal.color_rgb = dxf->border->diagonal.color_rgb;
                        c->cur_cf_rule.border->diagonal.has_color = dxf->border->diagonal.has_color;
                        c->cur_cf_rule.border->diagonal_up   = dxf->border->diagonal_up;
                        c->cur_cf_rule.border->diagonal_down = dxf->border->diagonal_down;
                    }
                }
            }
        }
        oxl_worksheet_add_cf_rule(c->ws, c->cur_cf_sqref, &c->cur_cf_rule);
        oxl_cf_rule_free_fields(&c->cur_cf_rule);
        memset(&c->cur_cf_rule, 0, sizeof(c->cur_cf_rule));
        c->state = SS_COND_FMT;
    }
    else if (strcmp(name, "conditionalFormatting") == 0 && c->state == SS_COND_FMT) {
        c->state = SS_NONE;
    }
}

static void XMLCALL sheet_char(void *ud, const char *s, int n) {
    SheetCtx *c = ud;
    if (c->state == SS_VALUE || c->state == SS_INLINE_T || c->state == SS_DV_FORMULA ||
        c->state == SS_CF_FORMULA) {
        cbuf_append(c, s, n);
    } else if (c->state == SS_FORMULA) {
        formula_buf_append(c, s, n);
    }
}

int oxl_parse_sheet(const char *buf, size_t len,
                    OxlWorksheet *ws, OxlWorkbook *wb,
                    const OxlHyperlinkRels *rels) {
    SheetCtx c;
    memset(&c, 0, sizeof(c));
    c.ws   = ws;
    c.wb   = wb;
    c.rels = rels;
    cbuf_reset(&c);
    formula_buf_reset(&c);

    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, &c);
    XML_SetElementHandler(p, sheet_start, sheet_end);
    XML_SetCharacterDataHandler(p, sheet_char);

    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);
    if (c.heap_buf) free(c.heap_buf);
    if (c.formula_heap) free(c.formula_heap);
    free(c.rich_buf);
    return (ok && !c.error) ? 0 : -1;
}
