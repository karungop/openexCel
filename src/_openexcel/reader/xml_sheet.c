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
}

static void XMLCALL sheet_char(void *ud, const char *s, int n) {
    SheetCtx *c = ud;
    if (c->state == SS_VALUE || c->state == SS_INLINE_T) {
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
