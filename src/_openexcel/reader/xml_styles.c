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
    enum {
        ST_NONE,
        ST_NUM_FMTS,
        ST_FONTS,       /* inside <fonts> */
        ST_FONT,        /* inside <fonts><font> */
        ST_FILLS,       /* inside <fills> */
        ST_FILL,        /* inside <fills><fill> */
        ST_BORDERS,     /* inside <borders> */
        ST_BORDER,      /* inside <borders><border> */
        ST_CELL_XFS,    /* inside <cellXfs> */
        ST_XF           /* inside <cellXfs><xf> (for <alignment> child) */
    } state;
    uint32_t    xf_index;
    int         error;

    /* Font being parsed */
    OxlFontDef   cur_font;

    /* Fill being parsed */
    OxlFillDef   cur_fill;

    /* Border being parsed */
    OxlBorderDef  cur_border;
    /* which side is being parsed currently (pointer into cur_border) */
    OxlBorderSide *cur_border_side;

    /* Per-XF data (for cellXfs entries) */
    uint32_t    xf_font_id;
    uint32_t    xf_fill_id;
    uint32_t    xf_border_id;
    OxlAlignDef xf_align;
    uint8_t     xf_apply_alignment;
} StylesCtx;

static char *safe_strdup(const char *s) { return s ? strdup(s) : NULL; }

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

    /* ── numFmts ── */
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

    /* ── fonts ── */
    } else if (strcmp(name, "fonts") == 0) {
        c->state = ST_FONTS;
    } else if (strcmp(name, "font") == 0 && c->state == ST_FONTS) {
        c->state = ST_FONT;
        memset(&c->cur_font, 0, sizeof(c->cur_font));
    } else if (c->state == ST_FONT) {
        if (strcmp(name, "b") == 0) {
            c->cur_font.bold = 1;
        } else if (strcmp(name, "i") == 0) {
            c->cur_font.italic = 1;
        } else if (strcmp(name, "u") == 0) {
            const char *val = attr(attrs, "val");
            if (val && strcmp(val, "double") == 0)
                c->cur_font.underline = 2;
            else
                c->cur_font.underline = 1;
        } else if (strcmp(name, "sz") == 0) {
            const char *val = attr(attrs, "val");
            if (val) c->cur_font.size = (float)atof(val);
        } else if (strcmp(name, "name") == 0) {
            const char *val = attr(attrs, "val");
            c->cur_font.name = safe_strdup(val);
        } else if (strcmp(name, "color") == 0) {
            const char *rgb = attr(attrs, "rgb");
            if (rgb) {
                c->cur_font.color_rgb = (uint32_t)strtoul(rgb, NULL, 16);
            } else {
                /* theme color — note it but don't set color_rgb from theme */
                const char *theme = attr(attrs, "theme");
                if (theme) c->cur_font.color_indexed = 1;
            }
        }

    /* ── fills ── */
    } else if (strcmp(name, "fills") == 0) {
        c->state = ST_FILLS;
    } else if (strcmp(name, "fill") == 0 && c->state == ST_FILLS) {
        c->state = ST_FILL;
        memset(&c->cur_fill, 0, sizeof(c->cur_fill));
    } else if (c->state == ST_FILL) {
        if (strcmp(name, "patternFill") == 0) {
            const char *pt = attr(attrs, "patternType");
            c->cur_fill.pattern_type = safe_strdup(pt ? pt : "none");
        } else if (strcmp(name, "fgColor") == 0) {
            const char *rgb = attr(attrs, "rgb");
            if (rgb) {
                c->cur_fill.fg_rgb = (uint32_t)strtoul(rgb, NULL, 16);
                c->cur_fill.fg_has_color = 1;
            }
        } else if (strcmp(name, "bgColor") == 0) {
            const char *rgb = attr(attrs, "rgb");
            if (rgb) {
                c->cur_fill.bg_rgb = (uint32_t)strtoul(rgb, NULL, 16);
                c->cur_fill.bg_has_color = 1;
            }
            /* ignore indexed bgColor — only rgb is useful */
        }

    /* ── borders ── */
    } else if (strcmp(name, "borders") == 0) {
        c->state = ST_BORDERS;
    } else if (strcmp(name, "border") == 0 && c->state == ST_BORDERS) {
        c->state = ST_BORDER;
        memset(&c->cur_border, 0, sizeof(c->cur_border));
        c->cur_border_side = NULL;
        /* diagonalUp / diagonalDown are attrs on <border> */
        const char *du = attr(attrs, "diagonalUp");
        const char *dd = attr(attrs, "diagonalDown");
        if (du && (strcmp(du, "1") == 0 || strcmp(du, "true") == 0))
            c->cur_border.diagonal_up = 1;
        if (dd && (strcmp(dd, "1") == 0 || strcmp(dd, "true") == 0))
            c->cur_border.diagonal_down = 1;
    } else if (c->state == ST_BORDER) {
        if (strcmp(name, "left") == 0) {
            c->cur_border_side = &c->cur_border.left;
            const char *style = attr(attrs, "style");
            if (style) c->cur_border_side->style = safe_strdup(style);
        } else if (strcmp(name, "right") == 0) {
            c->cur_border_side = &c->cur_border.right;
            const char *style = attr(attrs, "style");
            if (style) c->cur_border_side->style = safe_strdup(style);
        } else if (strcmp(name, "top") == 0) {
            c->cur_border_side = &c->cur_border.top;
            const char *style = attr(attrs, "style");
            if (style) c->cur_border_side->style = safe_strdup(style);
        } else if (strcmp(name, "bottom") == 0) {
            c->cur_border_side = &c->cur_border.bottom;
            const char *style = attr(attrs, "style");
            if (style) c->cur_border_side->style = safe_strdup(style);
        } else if (strcmp(name, "diagonal") == 0) {
            c->cur_border_side = &c->cur_border.diagonal;
            const char *style = attr(attrs, "style");
            if (style) c->cur_border_side->style = safe_strdup(style);
        } else if (strcmp(name, "color") == 0 && c->cur_border_side != NULL) {
            const char *rgb = attr(attrs, "rgb");
            if (rgb) {
                c->cur_border_side->color_rgb = (uint32_t)strtoul(rgb, NULL, 16);
                c->cur_border_side->has_color = 1;
            }
        }

    /* ── cellXfs ── */
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

        /* NEW: also read font/fill/border ids */
        const char *font_s   = attr(attrs, "fontId");
        const char *fill_s   = attr(attrs, "fillId");
        const char *border_s = attr(attrs, "borderId");
        c->xf_font_id   = font_s   ? (uint32_t)atoi(font_s)   : 0;
        c->xf_fill_id   = fill_s   ? (uint32_t)atoi(fill_s)   : 0;
        c->xf_border_id = border_s ? (uint32_t)atoi(border_s) : 0;
        memset(&c->xf_align, 0, sizeof(c->xf_align));
        c->xf_apply_alignment = 0;

        /* Transition to ST_XF so we can catch <alignment> child element */
        c->state = ST_XF;
    } else if (strcmp(name, "alignment") == 0 && c->state == ST_XF) {
        /* Parse alignment attributes */
        const char *horiz  = attr(attrs, "horizontal");
        const char *vert   = attr(attrs, "vertical");
        const char *wrap   = attr(attrs, "wrapText");
        const char *indent = attr(attrs, "indent");
        const char *rot    = attr(attrs, "textRotation");
        const char *shrink = attr(attrs, "shrinkToFit");
        if (horiz)  c->xf_align.horizontal    = safe_strdup(horiz);
        if (vert)   c->xf_align.vertical      = safe_strdup(vert);
        if (wrap && (strcmp(wrap, "1") == 0 || strcmp(wrap, "true") == 0))
            c->xf_align.wrap_text = 1;
        if (indent)  c->xf_align.indent        = atoi(indent);
        if (rot)     c->xf_align.text_rotation = atoi(rot);
        if (shrink && (strcmp(shrink, "1") == 0 || strcmp(shrink, "true") == 0))
            c->xf_align.shrink_to_fit = 1;
        c->xf_apply_alignment = 1;
    }
}

static void XMLCALL styles_end(void *ud, const char *name) {
    StylesCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "numFmts") == 0) {
        c->state = ST_NONE;
    } else if (strcmp(name, "font") == 0 && c->state == ST_FONT) {
        /* function declared in styles.h - implemented by styles.c */
        oxl_styles_add_font(c->styles, &c->cur_font);
        oxl_font_def_free_fields(&c->cur_font);
        c->state = ST_FONTS;
    } else if (strcmp(name, "fonts") == 0 && c->state == ST_FONTS) {
        c->state = ST_NONE;
    } else if (strcmp(name, "fill") == 0 && c->state == ST_FILL) {
        /* function declared in styles.h - implemented by styles.c */
        oxl_styles_add_fill(c->styles, &c->cur_fill);
        oxl_fill_def_free_fields(&c->cur_fill);
        c->state = ST_FILLS;
    } else if (strcmp(name, "fills") == 0 && c->state == ST_FILLS) {
        c->state = ST_NONE;
    } else if ((strcmp(name, "left") == 0 || strcmp(name, "right") == 0 ||
                strcmp(name, "top") == 0  || strcmp(name, "bottom") == 0 ||
                strcmp(name, "diagonal") == 0) && c->state == ST_BORDER) {
        c->cur_border_side = NULL;
    } else if (strcmp(name, "border") == 0 && c->state == ST_BORDER) {
        /* function declared in styles.h - implemented by styles.c */
        oxl_styles_add_border(c->styles, &c->cur_border);
        oxl_border_def_free_fields(&c->cur_border);
        c->state = ST_BORDERS;
    } else if (strcmp(name, "borders") == 0 && c->state == ST_BORDERS) {
        c->state = ST_NONE;
    } else if (strcmp(name, "xf") == 0 && c->state == ST_XF) {
        /* function declared in styles.h - implemented by styles.c */
        oxl_styles_set_xf_full(c->styles, c->xf_index,
                               c->xf_font_id, c->xf_fill_id, c->xf_border_id,
                               c->xf_apply_alignment ? &c->xf_align : NULL);
        oxl_align_def_free_fields(&c->xf_align);
        c->xf_index++;
        c->state = ST_CELL_XFS;
    } else if (strcmp(name, "cellXfs") == 0) {
        c->state = ST_NONE;
    }
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
