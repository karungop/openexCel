#include "xml_styles.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>

/* We build a map: numFmtId → is_date, then for each xf entry check its numFmtId. */

#define MAX_CUSTOM_FMTS 512

typedef struct {
    uint32_t id;
    int      is_date;
} NumFmtEntry;

typedef struct {
    OxlStyles  *styles;
    NumFmtEntry custom_fmts[MAX_CUSTOM_FMTS];
    uint32_t    custom_count;
    /* parse state */
    enum { ST_NONE, ST_NUM_FMTS, ST_CELL_XFS } state;
    uint32_t    xf_index;
    int         error;
} StylesCtx;

static const char *attr(const char **attrs, const char *key) {
    for (; *attrs; attrs += 2) if (strcmp(attrs[0], key) == 0) return attrs[1];
    return NULL;
}

static int custom_fmt_is_date(StylesCtx *c, uint32_t id) {
    for (uint32_t i = 0; i < c->custom_count; i++)
        if (c->custom_fmts[i].id == id) return c->custom_fmts[i].is_date;
    return 0;
}

static void XMLCALL styles_start(void *ud, const char *name, const char **attrs) {
    StylesCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "numFmts") == 0) {
        c->state = ST_NUM_FMTS;
    } else if (strcmp(name, "numFmt") == 0 && c->state == ST_NUM_FMTS) {
        const char *id_s  = attr(attrs, "numFmtId");
        const char *fmt_s = attr(attrs, "formatCode");
        if (id_s && fmt_s && c->custom_count < MAX_CUSTOM_FMTS) {
            uint32_t id = (uint32_t)atoi(id_s);
            c->custom_fmts[c->custom_count].id      = id;
            c->custom_fmts[c->custom_count].is_date = oxl_numfmt_str_is_date(fmt_s);
            c->custom_count++;
        }
    } else if (strcmp(name, "cellXfs") == 0) {
        c->state = ST_CELL_XFS;
        c->xf_index = 0;
    } else if (strcmp(name, "xf") == 0 && c->state == ST_CELL_XFS) {
        const char *id_s = attr(attrs, "numFmtId");
        int is_date = 0;
        if (id_s) {
            uint32_t id = (uint32_t)atoi(id_s);
            is_date = oxl_numfmt_id_is_date(id) || custom_fmt_is_date(c, id);
        }
        if (is_date) oxl_styles_set_date(c->styles, c->xf_index);
        else oxl_styles_resize(c->styles, c->xf_index + 1);
        c->xf_index++;
    }
}

static void XMLCALL styles_end(void *ud, const char *name) {
    StylesCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;
    if (strcmp(name, "numFmts") == 0 || strcmp(name, "cellXfs") == 0)
        c->state = ST_NONE;
}

int oxl_parse_styles(const char *buf, size_t len, OxlStyles *styles) {
    StylesCtx c = {0};
    c.styles = styles;

    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, &c);
    XML_SetElementHandler(p, styles_start, styles_end);

    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);
    return (ok && !c.error) ? 0 : -1;
}
