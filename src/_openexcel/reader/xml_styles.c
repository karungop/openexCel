#include "xml_styles.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

/* We build a map: numFmtId → format string + is_date, then for each xf entry
   check its numFmtId and record both the date flag and the fmt id. */

#define MAX_CUSTOM_FMTS 512
#define MAX_XF_COUNT    4096

typedef struct {
    uint32_t id;
    int      is_date;
    char     fmt_str[256];  /* truncated to 255 chars */
} FmtEntry;

typedef enum {
    ST_NONE,
    ST_NUM_FMTS,
    ST_FONTS,
    ST_FONT,
    ST_FILLS,
    ST_FILL,
    ST_BORDERS,
    ST_BORDER,
    ST_BORDER_SIDE,
    ST_CELL_XFS,
    ST_XF
} StyleState;

typedef struct {
    OxlStyles  *styles;
    FmtEntry    local_fmts[MAX_CUSTOM_FMTS];
    uint32_t    custom_count;
    /* parse state */
    StyleState  state;
    StyleState  prev_state;
    uint32_t    xf_index;
    int         error;
    /* current font being parsed */
    OxlFontDef  cur_font;
    /* current fill being parsed */
    OxlFillDef  cur_fill;
    /* current border being parsed */
    OxlBorderDef cur_border;
    /* which side of border we're currently in: 0=none,1=left,2=right,3=top,4=bottom,5=diag */
    int          border_side;
    /* current xf alignment */
    OxlAlignDef  cur_align;
    uint8_t      cur_xf_has_align;
    uint32_t     cur_xf_font_id;
    uint32_t     cur_xf_fill_id;
    uint32_t     cur_xf_border_id;
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

static uint32_t parse_color_attr(const char **attrs) {
    /* Try rgb="AARRGGBB" first */
    const char *rgb = attr(attrs, "rgb");
    if (rgb) {
        unsigned long v = strtoul(rgb, NULL, 16);
        return (uint32_t)v;
    }
    return 0;
}

static void XMLCALL styles_start(void *ud, const char *name, const char **attrs) {
    StylesCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    /* ---- numFmts ---- */
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
            size_t flen = strlen(fmt_s);
            if (flen > 255) flen = 255;
            memcpy(entry->fmt_str, fmt_s, flen);
            entry->fmt_str[flen] = '\0';
            c->custom_count++;
        }

    /* ---- fonts ---- */
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
            if (val && strcmp(val, "double") == 0) c->cur_font.underline = 2;
            else c->cur_font.underline = 1;
        } else if (strcmp(name, "sz") == 0) {
            const char *val = attr(attrs, "val");
            if (val) c->cur_font.size = (float)atof(val);
        } else if (strcmp(name, "name") == 0) {
            const char *val = attr(attrs, "val");
            if (val) { free(c->cur_font.name); c->cur_font.name = strdup(val); }
        } else if (strcmp(name, "color") == 0) {
            uint32_t rgb = parse_color_attr(attrs);
            if (rgb) c->cur_font.color_rgb = rgb;
        }

    /* ---- fills ---- */
    } else if (strcmp(name, "fills") == 0) {
        c->state = ST_FILLS;
    } else if (strcmp(name, "fill") == 0 && c->state == ST_FILLS) {
        c->state = ST_FILL;
        memset(&c->cur_fill, 0, sizeof(c->cur_fill));
    } else if (c->state == ST_FILL) {
        if (strcmp(name, "patternFill") == 0) {
            const char *pt = attr(attrs, "patternType");
            if (pt) { free(c->cur_fill.pattern_type); c->cur_fill.pattern_type = strdup(pt); }
        } else if (strcmp(name, "fgColor") == 0) {
            uint32_t rgb = parse_color_attr(attrs);
            if (rgb) { c->cur_fill.fg_rgb = rgb; c->cur_fill.fg_has_color = 1; }
        } else if (strcmp(name, "bgColor") == 0) {
            uint32_t rgb = parse_color_attr(attrs);
            if (rgb) { c->cur_fill.bg_rgb = rgb; c->cur_fill.bg_has_color = 1; }
        }

    /* ---- borders ---- */
    } else if (strcmp(name, "borders") == 0) {
        c->state = ST_BORDERS;
    } else if (strcmp(name, "border") == 0 && c->state == ST_BORDERS) {
        c->state = ST_BORDER;
        memset(&c->cur_border, 0, sizeof(c->cur_border));
        c->border_side = 0;
    } else if (c->state == ST_BORDER) {
        OxlBorderSide *side = NULL;
        if (strcmp(name, "left") == 0)     { c->border_side = 1; side = &c->cur_border.left; }
        else if (strcmp(name, "right") == 0)  { c->border_side = 2; side = &c->cur_border.right; }
        else if (strcmp(name, "top") == 0)    { c->border_side = 3; side = &c->cur_border.top; }
        else if (strcmp(name, "bottom") == 0) { c->border_side = 4; side = &c->cur_border.bottom; }
        else if (strcmp(name, "diagonal") == 0) { c->border_side = 5; side = &c->cur_border.diagonal; }
        if (side) {
            const char *st = attr(attrs, "style");
            if (st) { free(side->style); side->style = strdup(st); }
            c->state = ST_BORDER_SIDE;
        }
    } else if (c->state == ST_BORDER_SIDE && strcmp(name, "color") == 0) {
        OxlBorderSide *side = NULL;
        switch (c->border_side) {
            case 1: side = &c->cur_border.left;     break;
            case 2: side = &c->cur_border.right;    break;
            case 3: side = &c->cur_border.top;      break;
            case 4: side = &c->cur_border.bottom;   break;
            case 5: side = &c->cur_border.diagonal; break;
        }
        if (side) {
            uint32_t rgb = parse_color_attr(attrs);
            if (rgb) { side->color_rgb = rgb; side->has_color = 1; }
        }

    /* ---- cellXfs ---- */
    } else if (strcmp(name, "cellXfs") == 0) {
        c->state = ST_CELL_XFS;
        c->xf_index = 0;
    } else if (strcmp(name, "xf") == 0 && c->state == ST_CELL_XFS) {
        c->state = ST_XF;
        const char *id_s       = attr(attrs, "numFmtId");
        const char *font_id_s  = attr(attrs, "fontId");
        const char *fill_id_s  = attr(attrs, "fillId");
        const char *border_id_s = attr(attrs, "borderId");
        uint32_t id = 0;
        int is_date = 0;
        if (id_s) {
            id = (uint32_t)atoi(id_s);
            is_date = oxl_numfmt_id_is_date(id) || local_fmt_is_date(c, id);
        }
        if (is_date) oxl_styles_set_date(c->styles, c->xf_index);
        else oxl_styles_resize(c->styles, c->xf_index + 1);
        oxl_styles_set_xf_numfmt(c->styles, c->xf_index, id);

        c->cur_xf_font_id   = font_id_s   ? (uint32_t)atoi(font_id_s)   : 0;
        c->cur_xf_fill_id   = fill_id_s   ? (uint32_t)atoi(fill_id_s)   : 0;
        c->cur_xf_border_id = border_id_s ? (uint32_t)atoi(border_id_s) : 0;
        memset(&c->cur_align, 0, sizeof(c->cur_align));
        c->cur_xf_has_align = 0;

    } else if (c->state == ST_XF && strcmp(name, "alignment") == 0) {
        const char *horiz = attr(attrs, "horizontal");
        const char *vert  = attr(attrs, "vertical");
        const char *wrap  = attr(attrs, "wrapText");
        const char *shrink = attr(attrs, "shrinkToFit");
        const char *indent = attr(attrs, "indent");
        const char *textrot = attr(attrs, "textRotation");
        if (horiz)   { free(c->cur_align.horizontal); c->cur_align.horizontal = strdup(horiz); }
        if (vert)    { free(c->cur_align.vertical);   c->cur_align.vertical   = strdup(vert); }
        if (wrap)    c->cur_align.wrap_text     = (uint8_t)(atoi(wrap) ? 1 : 0);
        if (shrink)  c->cur_align.shrink_to_fit = (uint8_t)(atoi(shrink) ? 1 : 0);
        if (indent)  c->cur_align.indent        = (int32_t)atoi(indent);
        if (textrot) c->cur_align.text_rotation = (int32_t)atoi(textrot);
        c->cur_xf_has_align = 1;
    }
}

static void XMLCALL styles_end(void *ud, const char *name) {
    StylesCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "numFmts") == 0) {
        c->state = ST_NONE;
    } else if (strcmp(name, "font") == 0 && c->state == ST_FONT) {
        /* Add current font to registry */
        oxl_styles_add_font(c->styles, &c->cur_font);
        free(c->cur_font.name);
        memset(&c->cur_font, 0, sizeof(c->cur_font));
        c->state = ST_FONTS;
    } else if (strcmp(name, "fonts") == 0) {
        c->state = ST_NONE;
    } else if (strcmp(name, "fill") == 0 && c->state == ST_FILL) {
        oxl_styles_add_fill(c->styles, &c->cur_fill);
        free(c->cur_fill.pattern_type);
        memset(&c->cur_fill, 0, sizeof(c->cur_fill));
        c->state = ST_FILLS;
    } else if (strcmp(name, "fills") == 0) {
        c->state = ST_NONE;
    } else if ((strcmp(name, "left") == 0 || strcmp(name, "right") == 0 ||
                strcmp(name, "top") == 0 || strcmp(name, "bottom") == 0 ||
                strcmp(name, "diagonal") == 0) && c->state == ST_BORDER_SIDE) {
        c->state = ST_BORDER;
        c->border_side = 0;
    } else if (strcmp(name, "border") == 0 && c->state == ST_BORDER) {
        oxl_styles_add_border(c->styles, &c->cur_border);
        /* Free side styles */
        free(c->cur_border.left.style);
        free(c->cur_border.right.style);
        free(c->cur_border.top.style);
        free(c->cur_border.bottom.style);
        free(c->cur_border.diagonal.style);
        memset(&c->cur_border, 0, sizeof(c->cur_border));
        c->state = ST_BORDERS;
    } else if (strcmp(name, "borders") == 0) {
        c->state = ST_NONE;
    } else if (strcmp(name, "xf") == 0 && c->state == ST_XF) {
        /* Store full XF info */
        oxl_styles_set_xf_full(c->styles, c->xf_index,
                                c->cur_xf_font_id, c->cur_xf_fill_id, c->cur_xf_border_id,
                                c->cur_xf_has_align ? &c->cur_align : NULL);
        /* Free alignment strings */
        free(c->cur_align.horizontal);
        free(c->cur_align.vertical);
        memset(&c->cur_align, 0, sizeof(c->cur_align));
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

    /* Cleanup in case of error */
    free(c.cur_font.name);
    free(c.cur_fill.pattern_type);
    free(c.cur_border.left.style);
    free(c.cur_border.right.style);
    free(c.cur_border.top.style);
    free(c.cur_border.bottom.style);
    free(c.cur_border.diagonal.style);
    free(c.cur_align.horizontal);
    free(c.cur_align.vertical);
    return -1;
}
