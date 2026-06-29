#include "xml_styles.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>

/* We build a map: numFmtId → format string + is_date, then for each xf entry
   check its numFmtId and record both the date flag and the fmt id. */

#define MAX_CUSTOM_FMTS 512
#define MAX_XF_COUNT    4096

typedef struct {
    uint32_t id;
    int      is_date;
    char     fmt_str[256];  /* truncated to 255 chars */
} FmtEntry;

typedef struct {
    OxlStyles  *styles;
    FmtEntry    local_fmts[MAX_CUSTOM_FMTS];
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

static int local_fmt_is_date(StylesCtx *c, uint32_t id) {
    for (uint32_t i = 0; i < c->custom_count; i++)
        if (c->local_fmts[i].id == id) return c->local_fmts[i].is_date;
    return 0;
}

static const char *local_fmt_str(StylesCtx *c, uint32_t id) {
    for (uint32_t i = 0; i < c->custom_count; i++)
        if (c->local_fmts[i].id == id) return c->local_fmts[i].fmt_str;
    return NULL;
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
            FmtEntry *entry = &c->local_fmts[c->custom_count];
            entry->id      = id;
            entry->is_date = oxl_numfmt_str_is_date(fmt_s);
            /* Copy format string, truncate to 255 chars */
            size_t flen = strlen(fmt_s);
            if (flen > 255) flen = 255;
            memcpy(entry->fmt_str, fmt_s, flen);
            entry->fmt_str[flen] = '\0';
            c->custom_count++;
        }
    } else if (strcmp(name, "cellXfs") == 0) {
        c->state = ST_CELL_XFS;
        c->xf_index = 0;
    } else if (strcmp(name, "xf") == 0 && c->state == ST_CELL_XFS) {
        const char *id_s = attr(attrs, "numFmtId");
        uint32_t id = 0;
        int is_date = 0;
        if (id_s) {
            id = (uint32_t)atoi(id_s);
            is_date = oxl_numfmt_id_is_date(id) || local_fmt_is_date(c, id);
        }
        if (is_date) oxl_styles_set_date(c->styles, c->xf_index);
        else oxl_styles_resize(c->styles, c->xf_index + 1);

        /* Record the numFmtId for this xf */
        oxl_styles_set_xf_numfmt(c->styles, c->xf_index, id);

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

    if (ok && !c.error) {
        /* Populate styles->custom_fmts from local_fmts (custom ids >= 164) */
        for (uint32_t i = 0; i < c.custom_count; i++) {
            if (c.local_fmts[i].id >= 164) {
                oxl_styles_add_custom_fmt(styles, c.local_fmts[i].id,
                                          c.local_fmts[i].fmt_str);
            }
        }
        return 0;
    }
    return -1;
}
