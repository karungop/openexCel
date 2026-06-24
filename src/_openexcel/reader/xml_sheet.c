#include "xml_sheet.h"
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
    /* rich inline string accumulator */
    char         *rich_buf;
    size_t        rich_len;
    size_t        rich_cap;
    int           error;
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
    if (c->cbuf_len == 0 && !c->cur_is_inline) return;

    OxlCell cell;
    memset(&cell, 0, sizeof(cell));
    cell.row       = c->cur_row;
    cell.col       = c->cur_col;
    cell.style_idx = c->cur_style;

    /* null-terminate cbuf */
    if (c->cbuf_len < c->cbuf_cap) c->cbuf[c->cbuf_len] = '\0';

    if (c->cur_is_inline) {
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
        if (strcmp(name, "sheetData") == 0) c->state = SS_SHEET_DATA;
        break;
    case SS_SHEET_DATA:
        if (strcmp(name, "row") == 0) {
            const char *r = attr(attrs, "r");
            c->cur_row = r ? (uint32_t)(atoi(r) - 1) : c->cur_row;
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
        cbuf_reset(c);
        c->state = SS_CELL;
    } else if (strcmp(name, "t") == 0 && c->state == SS_INLINE_T) {
        rich_append(c, c->cbuf, c->cbuf_len);
        c->state = SS_INLINE_SI;
    } else if (strcmp(name, "is") == 0 && c->state == SS_INLINE_SI) {
        c->state = SS_CELL;
    } else if (strcmp(name, "c") == 0 && c->state == SS_CELL) {
        emit_cell(c);
        c->state = SS_ROW;
    } else if (strcmp(name, "row") == 0 && c->state == SS_ROW) {
        c->state = SS_SHEET_DATA;
    } else if (strcmp(name, "sheetData") == 0 && c->state == SS_SHEET_DATA) {
        c->state = SS_NONE;
    }
}

static void XMLCALL sheet_char(void *ud, const char *s, int n) {
    SheetCtx *c = ud;
    if (c->state == SS_VALUE || c->state == SS_INLINE_T) {
        cbuf_append(c, s, n);
    }
}

int oxl_parse_sheet(const char *buf, size_t len,
                    OxlWorksheet *ws, OxlWorkbook *wb) {
    SheetCtx c;
    memset(&c, 0, sizeof(c));
    c.ws   = ws;
    c.wb   = wb;
    cbuf_reset(&c);

    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, &c);
    XML_SetElementHandler(p, sheet_start, sheet_end);
    XML_SetCharacterDataHandler(p, sheet_char);

    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);
    if (c.heap_buf) free(c.heap_buf);
    free(c.rich_buf);
    return (ok && !c.error) ? 0 : -1;
}
