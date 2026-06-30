#include "styles.h"
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <ctype.h>

/* Built-in Excel numFmt IDs that are date/time formats.
   Reference: ECMA-376 Part 1, §18.8.30 */
static const uint32_t BUILTIN_DATE_FMTS[] = {
    14, 15, 16, 17,          /* m/d/yy, d-mmm-yy, d-mmm, mmm-yy */
    18, 19, 20, 21, 22,      /* h:mm AM/PM, h:mm:ss AM/PM, h:mm, h:mm:ss, m/d/yy h:mm */
    45, 46, 47,              /* mm:ss, [h]:mm:ss, mmss.0 */
};

/* Built-in format string table */
static const struct { uint32_t id; const char *str; } BUILTIN_FMTS[] = {
    {1,  "0"},
    {2,  "0.00"},
    {3,  "#,##0"},
    {4,  "#,##0.00"},
    {9,  "0%"},
    {10, "0.00%"},
    {11, "0.00E+00"},
    {12, "# ?/?"},
    {13, "# ??/??"},
    {14, "m/d/yyyy"},
    {15, "d-mmm-yy"},
    {16, "d-mmm"},
    {17, "mmm-yy"},
    {18, "h:mm AM/PM"},
    {19, "h:mm:ss AM/PM"},
    {20, "h:mm"},
    {21, "h:mm:ss"},
    {22, "m/d/yy h:mm"},
    {37, "#,##0 ;(#,##0)"},
    {38, "#,##0 ;[Red](#,##0)"},
    {39, "#,##0.00;(#,##0.00)"},
    {40, "#,##0.00;[Red](#,##0.00)"},
    {45, "mm:ss"},
    {46, "[h]:mm:ss"},
    {47, "mmss.0"},
    {48, "##0.0E+0"},
    {49, "@"},
};
#define BUILTIN_FMTS_COUNT (sizeof(BUILTIN_FMTS)/sizeof(BUILTIN_FMTS[0]))

int oxl_numfmt_id_is_date(uint32_t id) {
    for (size_t i = 0; i < sizeof(BUILTIN_DATE_FMTS)/sizeof(BUILTIN_DATE_FMTS[0]); i++) {
        if (BUILTIN_DATE_FMTS[i] == id) return 1;
    }
    return 0;
}

int oxl_numfmt_str_is_date(const char *fmt) {
    if (!fmt) return 0;
    /* Skip quoted substrings (e.g. "text") and brackets ([red]) */
    int in_quote  = 0;
    int in_bracket = 0;
    for (const char *p = fmt; *p; p++) {
        if (*p == '"') { in_quote = !in_quote; continue; }
        if (in_quote) continue;
        if (*p == '[') { in_bracket++; continue; }
        if (*p == ']') { if (in_bracket > 0) in_bracket--; continue; }
        if (in_bracket) continue;
        char c = (char)tolower((unsigned char)*p);
        if (c == 'y' || c == 'd') return 1;
        /* 'm' is month only if not preceded by '[h]' (hours in elapsed-time fmt) */
        if (c == 'm') return 1;
        /* 'h' alone means hours (time), combined with date chars makes it datetime */
        if (c == 'h') return 1;
    }
    return 0;
}

void oxl_styles_init(OxlStyles *s) {
    s->date_xf_bits = NULL;
    s->xf_count = 0;
    s->xf_num_fmt_ids = NULL;
    s->custom_fmts = NULL;
    s->custom_fmt_count = 0;
    s->custom_fmt_cap = 0;
    s->xf_registry = NULL;
    s->xf_registry_count = 0;
    s->xf_registry_cap = 0;
    s->fonts = NULL;
    s->font_count = 0;
    s->font_cap = 0;
    s->fills = NULL;
    s->fill_count = 0;
    s->fill_cap = 0;
    s->borders = NULL;
    s->border_count = 0;
    s->border_cap = 0;
}

void oxl_font_def_free_fields(OxlFontDef *f) {
    free(f->name); f->name = NULL;
}

void oxl_fill_def_free_fields(OxlFillDef *f) {
    free(f->pattern_type); f->pattern_type = NULL;
}

void oxl_border_side_free_fields(OxlBorderSide *side) {
    free(side->style); side->style = NULL;
}

void oxl_border_def_free_fields(OxlBorderDef *b) {
    oxl_border_side_free_fields(&b->left);
    oxl_border_side_free_fields(&b->right);
    oxl_border_side_free_fields(&b->top);
    oxl_border_side_free_fields(&b->bottom);
    oxl_border_side_free_fields(&b->diagonal);
}

void oxl_align_def_free_fields(OxlAlignDef *a) {
    free(a->horizontal); a->horizontal = NULL;
    free(a->vertical);   a->vertical   = NULL;
}

void oxl_styles_free(OxlStyles *s) {
    free(s->date_xf_bits);
    s->date_xf_bits = NULL;

    free(s->xf_num_fmt_ids);
    s->xf_num_fmt_ids = NULL;

    if (s->custom_fmts) {
        for (uint32_t i = 0; i < s->custom_fmt_count; i++)
            free(s->custom_fmts[i].fmt_str);
        free(s->custom_fmts);
        s->custom_fmts = NULL;
    }
    s->custom_fmt_count = 0;
    s->custom_fmt_cap = 0;

    if (s->xf_registry) {
        for (uint32_t i = 0; i < s->xf_registry_count; i++) {
            free(s->xf_registry[i].fmt_str);
            oxl_align_def_free_fields(&s->xf_registry[i].align);
        }
        free(s->xf_registry);
        s->xf_registry = NULL;
    }
    s->xf_registry_count = 0;
    s->xf_registry_cap = 0;

    if (s->fonts) {
        for (uint32_t i = 0; i < s->font_count; i++)
            oxl_font_def_free_fields(&s->fonts[i]);
        free(s->fonts);
        s->fonts = NULL;
    }
    s->font_count = 0;
    s->font_cap = 0;

    if (s->fills) {
        for (uint32_t i = 0; i < s->fill_count; i++)
            oxl_fill_def_free_fields(&s->fills[i]);
        free(s->fills);
        s->fills = NULL;
    }
    s->fill_count = 0;
    s->fill_cap = 0;

    if (s->borders) {
        for (uint32_t i = 0; i < s->border_count; i++)
            oxl_border_def_free_fields(&s->borders[i]);
        free(s->borders);
        s->borders = NULL;
    }
    s->border_count = 0;
    s->border_cap = 0;

    s->xf_count = 0;
}

/* ---- Read-side: add font/fill/border from XML parser ---- */

uint32_t oxl_styles_add_font(OxlStyles *s, const OxlFontDef *font) {
    if (s->font_count >= s->font_cap) {
        uint32_t new_cap = s->font_cap ? s->font_cap * 2 : 4;
        OxlFontDef *p = realloc(s->fonts, new_cap * sizeof(*p));
        if (!p) return 0;
        s->fonts = p;
        s->font_cap = new_cap;
    }
    OxlFontDef *dst = &s->fonts[s->font_count];
    *dst = *font;
    dst->name = font->name ? strdup(font->name) : NULL;
    return s->font_count++;
}

uint32_t oxl_styles_add_fill(OxlStyles *s, const OxlFillDef *fill) {
    if (s->fill_count >= s->fill_cap) {
        uint32_t new_cap = s->fill_cap ? s->fill_cap * 2 : 4;
        OxlFillDef *p = realloc(s->fills, new_cap * sizeof(*p));
        if (!p) return 0;
        s->fills = p;
        s->fill_cap = new_cap;
    }
    OxlFillDef *dst = &s->fills[s->fill_count];
    *dst = *fill;
    dst->pattern_type = fill->pattern_type ? strdup(fill->pattern_type) : NULL;
    return s->fill_count++;
}

uint32_t oxl_styles_add_border(OxlStyles *s, const OxlBorderDef *border) {
    if (s->border_count >= s->border_cap) {
        uint32_t new_cap = s->border_cap ? s->border_cap * 2 : 4;
        OxlBorderDef *p = realloc(s->borders, new_cap * sizeof(*p));
        if (!p) return 0;
        s->borders = p;
        s->border_cap = new_cap;
    }
    OxlBorderDef *dst = &s->borders[s->border_count];
    *dst = *border;
    /* Deep copy side styles */
#define COPY_SIDE(side) dst->side.style = border->side.style ? strdup(border->side.style) : NULL
    COPY_SIDE(left); COPY_SIDE(right); COPY_SIDE(top); COPY_SIDE(bottom); COPY_SIDE(diagonal);
#undef COPY_SIDE
    return s->border_count++;
}

/* Read-side: set full XF info (font/fill/border/align) for an xf_index */
void oxl_styles_set_xf_full(OxlStyles *s, uint32_t xf_index,
                              uint32_t font_id, uint32_t fill_id,
                              uint32_t border_id, const OxlAlignDef *align) {
    /* Ensure xf_registry is large enough */
    while (s->xf_registry_count <= xf_index) {
        if (s->xf_registry_count >= s->xf_registry_cap) {
            uint32_t new_cap = s->xf_registry_cap ? s->xf_registry_cap * 2 : 4;
            void *p = realloc(s->xf_registry, new_cap * sizeof(*s->xf_registry));
            if (!p) return;
            s->xf_registry = p;
            s->xf_registry_cap = new_cap;
        }
        memset(&s->xf_registry[s->xf_registry_count], 0, sizeof(*s->xf_registry));
        s->xf_registry_count++;
    }
    OxlXfRecord *xf = &s->xf_registry[xf_index];
    xf->font_id   = font_id;
    xf->fill_id   = fill_id;
    xf->border_id = border_id;

    /* Also sync num_fmt_id and fmt_str from xf_num_fmt_ids if available */
    if (s->xf_num_fmt_ids && xf_index < s->xf_count) {
        xf->num_fmt_id = s->xf_num_fmt_ids[xf_index];
        if (!xf->fmt_str) {
            const char *fs = oxl_styles_get_numfmt_str(s, (uint16_t)xf_index);
            xf->fmt_str = strdup(fs ? fs : "General");
        }
    }

    if (align) {
        oxl_align_def_free_fields(&xf->align);
        xf->align.horizontal    = align->horizontal    ? strdup(align->horizontal)    : NULL;
        xf->align.vertical      = align->vertical      ? strdup(align->vertical)      : NULL;
        xf->align.indent        = align->indent;
        xf->align.text_rotation = align->text_rotation;
        xf->align.wrap_text     = align->wrap_text;
        xf->align.shrink_to_fit = align->shrink_to_fit;
        xf->apply_alignment = (align->horizontal || align->vertical ||
                               align->wrap_text || align->shrink_to_fit ||
                               align->indent || align->text_rotation) ? 1 : 0;
    }
}

/* ---- Accessors ---- */

const OxlXfRecord *oxl_styles_get_xf(const OxlStyles *s, uint16_t xf_index) {
    if (!s->xf_registry || xf_index >= s->xf_registry_count) return NULL;
    return &s->xf_registry[xf_index];
}

const OxlFontDef *oxl_styles_get_font(const OxlStyles *s, uint32_t font_id) {
    if (!s->fonts || font_id >= s->font_count) return NULL;
    return &s->fonts[font_id];
}

const OxlFillDef *oxl_styles_get_fill(const OxlStyles *s, uint32_t fill_id) {
    if (!s->fills || fill_id >= s->fill_count) return NULL;
    return &s->fills[fill_id];
}

const OxlBorderDef *oxl_styles_get_border(const OxlStyles *s, uint32_t border_id) {
    if (!s->borders || border_id >= s->border_count) return NULL;
    return &s->borders[border_id];
}

/* ---- Write-side: find-or-create with dedup ---- */

static int fonts_equal(const OxlFontDef *a, const OxlFontDef *b) {
    if (a->bold != b->bold || a->italic != b->italic || a->underline != b->underline) return 0;
    if (a->color_rgb != b->color_rgb) return 0;
    if (a->size != b->size) return 0;
    if ((a->name == NULL) != (b->name == NULL)) return 0;
    if (a->name && b->name && strcmp(a->name, b->name) != 0) return 0;
    return 1;
}

uint32_t oxl_styles_get_or_add_font(OxlStyles *s, const OxlFontDef *font) {
    for (uint32_t i = 0; i < s->font_count; i++) {
        if (fonts_equal(&s->fonts[i], font)) return i;
    }
    return oxl_styles_add_font(s, font);
}

static int fills_equal(const OxlFillDef *a, const OxlFillDef *b) {
    if (a->fg_rgb != b->fg_rgb || a->bg_rgb != b->bg_rgb) return 0;
    if (a->fg_has_color != b->fg_has_color || a->bg_has_color != b->bg_has_color) return 0;
    if ((a->pattern_type == NULL) != (b->pattern_type == NULL)) return 0;
    if (a->pattern_type && b->pattern_type && strcmp(a->pattern_type, b->pattern_type) != 0) return 0;
    return 1;
}

uint32_t oxl_styles_get_or_add_fill(OxlStyles *s, const OxlFillDef *fill) {
    for (uint32_t i = 0; i < s->fill_count; i++) {
        if (fills_equal(&s->fills[i], fill)) return i;
    }
    return oxl_styles_add_fill(s, fill);
}

static int sides_equal(const OxlBorderSide *a, const OxlBorderSide *b) {
    if (a->color_rgb != b->color_rgb || a->has_color != b->has_color) return 0;
    if ((a->style == NULL) != (b->style == NULL)) return 0;
    if (a->style && b->style && strcmp(a->style, b->style) != 0) return 0;
    return 1;
}

static int borders_equal(const OxlBorderDef *a, const OxlBorderDef *b) {
    return sides_equal(&a->left, &b->left) &&
           sides_equal(&a->right, &b->right) &&
           sides_equal(&a->top, &b->top) &&
           sides_equal(&a->bottom, &b->bottom);
}

uint32_t oxl_styles_get_or_add_border(OxlStyles *s, const OxlBorderDef *border) {
    for (uint32_t i = 0; i < s->border_count; i++) {
        if (borders_equal(&s->borders[i], border)) return i;
    }
    return oxl_styles_add_border(s, border);
}

static int aligns_equal(const OxlAlignDef *a, const OxlAlignDef *b) {
    if (a->indent != b->indent || a->text_rotation != b->text_rotation) return 0;
    if (a->wrap_text != b->wrap_text || a->shrink_to_fit != b->shrink_to_fit) return 0;
    if ((a->horizontal == NULL) != (b->horizontal == NULL)) return 0;
    if (a->horizontal && b->horizontal && strcmp(a->horizontal, b->horizontal) != 0) return 0;
    if ((a->vertical == NULL) != (b->vertical == NULL)) return 0;
    if (a->vertical && b->vertical && strcmp(a->vertical, b->vertical) != 0) return 0;
    return 1;
}

uint16_t oxl_styles_get_or_add_xf_full(OxlStyles *s, const char *fmt_str,
                                         uint32_t font_id, uint32_t fill_id,
                                         uint32_t border_id, const OxlAlignDef *align) {
    uint8_t has_align = align && (align->horizontal || align->vertical ||
                                   align->wrap_text || align->shrink_to_fit ||
                                   align->indent || align->text_rotation);

    /* Search for existing matching XF */
    for (uint32_t i = 0; i < s->xf_registry_count; i++) {
        OxlXfRecord *xf = &s->xf_registry[i];
        if (xf->font_id != font_id || xf->fill_id != fill_id || xf->border_id != border_id) continue;
        if (xf->apply_alignment != (has_align ? 1 : 0)) continue;
        /* Check fmt_str */
        const char *xfmt = xf->fmt_str ? xf->fmt_str : "General";
        const char *needle = fmt_str ? fmt_str : "General";
        if (strcmp(xfmt, needle) != 0) continue;
        /* Check alignment */
        if (has_align && !aligns_equal(&xf->align, align)) continue;
        return (uint16_t)i;
    }

    /* Not found: create */
    if (s->xf_registry_count >= s->xf_registry_cap) {
        uint32_t new_cap = s->xf_registry_cap ? s->xf_registry_cap * 2 : 4;
        void *p = realloc(s->xf_registry, new_cap * sizeof(*s->xf_registry));
        if (!p) return 0;
        s->xf_registry = p;
        s->xf_registry_cap = new_cap;
    }

    uint32_t new_idx = s->xf_registry_count;
    OxlXfRecord *xf = &s->xf_registry[new_idx];
    memset(xf, 0, sizeof(*xf));

    /* Determine numFmtId and custom fmt */
    const char *use_fmt = (fmt_str && *fmt_str) ? fmt_str : "General";
    int is_general = (strcmp(use_fmt, "General") == 0);
    uint32_t num_fmt_id = 0;

    if (!is_general) {
        int found_builtin = 0;
        for (size_t i = 0; i < BUILTIN_FMTS_COUNT; i++) {
            if (strcasecmp(BUILTIN_FMTS[i].str, use_fmt) == 0) {
                num_fmt_id = BUILTIN_FMTS[i].id;
                found_builtin = 1;
                break;
            }
        }
        if (!found_builtin) {
            /* Check if already added as custom */
            int found_custom = 0;
            for (uint32_t i = 0; i < s->custom_fmt_count; i++) {
                if (strcmp(s->custom_fmts[i].fmt_str, use_fmt) == 0) {
                    num_fmt_id = s->custom_fmts[i].id;
                    found_custom = 1;
                    break;
                }
            }
            if (!found_custom) {
                num_fmt_id = 164 + s->custom_fmt_count;
                oxl_styles_add_custom_fmt(s, num_fmt_id, use_fmt);
            }
        }
    }

    xf->fmt_str   = strdup(use_fmt);
    xf->num_fmt_id = num_fmt_id;
    xf->font_id   = font_id;
    xf->fill_id   = fill_id;
    xf->border_id = border_id;

    if (align) {
        xf->align.horizontal    = align->horizontal    ? strdup(align->horizontal)    : NULL;
        xf->align.vertical      = align->vertical      ? strdup(align->vertical)      : NULL;
        xf->align.indent        = align->indent;
        xf->align.text_rotation = align->text_rotation;
        xf->align.wrap_text     = align->wrap_text;
        xf->align.shrink_to_fit = align->shrink_to_fit;
        xf->apply_alignment     = has_align ? 1 : 0;
    }

    s->xf_registry_count++;

    /* Sync xf_count / date bits */
    oxl_styles_resize(s, new_idx + 1);
    if (s->xf_num_fmt_ids) s->xf_num_fmt_ids[new_idx] = num_fmt_id;
    if (!is_general && oxl_numfmt_str_is_date(use_fmt))
        oxl_styles_set_date(s, new_idx);

    return (uint16_t)new_idx;
}

void oxl_styles_resize(OxlStyles *s, uint32_t new_count) {
    uint32_t bytes = (new_count + 7) / 8;
    uint8_t *p = realloc(s->date_xf_bits, bytes);
    if (!p) return;
    /* Zero out newly added bytes */
    if (new_count > s->xf_count) {
        uint32_t old_bytes = (s->xf_count + 7) / 8;
        memset(p + old_bytes, 0, bytes - old_bytes);
    }
    s->date_xf_bits = p;

    /* Also resize xf_num_fmt_ids */
    if (new_count > s->xf_count) {
        uint32_t *fp = realloc(s->xf_num_fmt_ids, new_count * sizeof(uint32_t));
        if (fp) {
            /* Zero-fill new slots */
            memset(fp + s->xf_count, 0, (new_count - s->xf_count) * sizeof(uint32_t));
            s->xf_num_fmt_ids = fp;
        }
    }

    s->xf_count = new_count;
}

void oxl_styles_set_date(OxlStyles *s, uint32_t xf_index) {
    if (xf_index >= s->xf_count) oxl_styles_resize(s, xf_index + 1);
    s->date_xf_bits[xf_index / 8] |= (uint8_t)(1u << (xf_index % 8));
}

int oxl_styles_is_date(const OxlStyles *s, uint16_t xf_index) {
    if (!s->date_xf_bits || xf_index >= s->xf_count) return 0;
    return (s->date_xf_bits[xf_index / 8] >> (xf_index % 8)) & 1;
}

/* Read-side: set the numFmtId for an xf_index, growing arrays if needed */
void oxl_styles_set_xf_numfmt(OxlStyles *s, uint32_t xf_index, uint32_t num_fmt_id) {
    if (xf_index >= s->xf_count) {
        oxl_styles_resize(s, xf_index + 1);
    }
    if (s->xf_num_fmt_ids) {
        s->xf_num_fmt_ids[xf_index] = num_fmt_id;
    }
}

/* Read-side: add a custom numFmt (id >= 164) with its format string */
void oxl_styles_add_custom_fmt(OxlStyles *s, uint32_t id, const char *fmt_str) {
    if (s->custom_fmt_count >= s->custom_fmt_cap) {
        uint32_t new_cap = s->custom_fmt_cap ? s->custom_fmt_cap * 2 : 8;
        void *p = realloc(s->custom_fmts, new_cap * sizeof(*s->custom_fmts));
        if (!p) return;
        s->custom_fmts = p;
        s->custom_fmt_cap = new_cap;
    }
    s->custom_fmts[s->custom_fmt_count].id = id;
    s->custom_fmts[s->custom_fmt_count].fmt_str = strdup(fmt_str ? fmt_str : "");
    s->custom_fmt_count++;
}

/* Read-side: get the format string for an xf index */
const char *oxl_styles_get_numfmt_str(const OxlStyles *s, uint16_t xf_index) {
    if (!s->xf_num_fmt_ids || xf_index >= s->xf_count) return NULL;
    uint32_t id = s->xf_num_fmt_ids[xf_index];
    if (id == 0) return NULL;  /* General */

    if (id < 164) {
        /* Search built-in table */
        for (size_t i = 0; i < BUILTIN_FMTS_COUNT; i++) {
            if (BUILTIN_FMTS[i].id == id) return BUILTIN_FMTS[i].str;
        }
        return NULL;
    }

    /* Search custom fmts */
    for (uint32_t i = 0; i < s->custom_fmt_count; i++) {
        if (s->custom_fmts[i].id == id) return s->custom_fmts[i].fmt_str;
    }
    return NULL;
}

/* Write-side: find or create an XF entry for a format string */
uint16_t oxl_styles_get_or_add_xf(OxlStyles *s, const char *fmt_str) {
    /* NULL or empty → General (xf 0) */
    if (!fmt_str || fmt_str[0] == '\0') return 0;

    /* Search existing registry */
    for (uint32_t i = 0; i < s->xf_registry_count; i++) {
        if (s->xf_registry[i].fmt_str && strcmp(s->xf_registry[i].fmt_str, fmt_str) == 0)
            return (uint16_t)i;
    }

    /* Not found: determine numFmtId */
    uint32_t num_fmt_id = 0;
    int is_general = (strcmp(fmt_str, "General") == 0);

    if (!is_general) {
        /* Check built-in table for matching format string */
        int found_builtin = 0;
        for (size_t i = 0; i < BUILTIN_FMTS_COUNT; i++) {
            if (strcasecmp(BUILTIN_FMTS[i].str, fmt_str) == 0) {
                num_fmt_id = BUILTIN_FMTS[i].id;
                found_builtin = 1;
                break;
            }
        }

        if (!found_builtin) {
            /* Custom: id = 164 + count of current custom fmts */
            num_fmt_id = 164 + s->custom_fmt_count;
            /* Add to custom_fmts */
            oxl_styles_add_custom_fmt(s, num_fmt_id, fmt_str);
        }
    }
    /* is_general → num_fmt_id stays 0 */

    /* Grow xf_registry if needed */
    if (s->xf_registry_count >= s->xf_registry_cap) {
        uint32_t new_cap = s->xf_registry_cap ? s->xf_registry_cap * 2 : 4;
        void *p = realloc(s->xf_registry, new_cap * sizeof(*s->xf_registry));
        if (!p) return 0;
        s->xf_registry = p;
        s->xf_registry_cap = new_cap;
    }

    uint32_t new_idx = s->xf_registry_count;
    memset(&s->xf_registry[new_idx], 0, sizeof(s->xf_registry[new_idx]));
    s->xf_registry[new_idx].fmt_str = strdup(fmt_str);
    s->xf_registry[new_idx].num_fmt_id = num_fmt_id;
    s->xf_registry_count++;

    /* Keep xf_count, date_xf_bits, and xf_num_fmt_ids in sync */
    oxl_styles_resize(s, new_idx + 1);
    if (s->xf_num_fmt_ids) s->xf_num_fmt_ids[new_idx] = num_fmt_id;

    /* Mark as date if applicable */
    if (!is_general && oxl_numfmt_str_is_date(fmt_str)) {
        oxl_styles_set_date(s, new_idx);
    }

    return (uint16_t)new_idx;
}

/* Initialize write-side defaults (xf 0 = General, xf 1 = date YYYY-MM-DD) */
void oxl_styles_init_write_defaults(OxlStyles *s) {
    /* Ensure default font (id=0) exists */
    if (s->font_count == 0) {
        OxlFontDef default_font = {0};
        default_font.name = "Calibri";
        default_font.size = 11.0f;
        oxl_styles_add_font(s, &default_font);
    }

    /* Ensure default fills (id=0 = none, id=1 = gray125) exist */
    if (s->fill_count == 0) {
        OxlFillDef fill0 = {0}; fill0.pattern_type = "none";
        oxl_styles_add_fill(s, &fill0);
        OxlFillDef fill1 = {0}; fill1.pattern_type = "gray125";
        oxl_styles_add_fill(s, &fill1);
    }

    /* Ensure default border (id=0) exists */
    if (s->border_count == 0) {
        OxlBorderDef border0;
        memset(&border0, 0, sizeof(border0));
        oxl_styles_add_border(s, &border0);
    }

    /* If xf_registry was already populated by oxl_styles_set_xf_full during parse,
       just ensure fmt_str fields are set and we have at least 2 entries */
    if (s->xf_registry_count > 0) {
        /* Patch any missing fmt_str entries */
        for (uint32_t i = 0; i < s->xf_registry_count; i++) {
            if (!s->xf_registry[i].fmt_str) {
                const char *fs = oxl_styles_get_numfmt_str(s, (uint16_t)i);
                s->xf_registry[i].fmt_str = strdup(fs ? fs : "General");
                if (s->xf_num_fmt_ids && i < s->xf_count)
                    s->xf_registry[i].num_fmt_id = s->xf_num_fmt_ids[i];
            }
        }
        if (s->xf_registry_count < 2) {
            oxl_styles_get_or_add_xf(s, "YYYY-MM-DD");
        }
        return;
    }

    if (s->xf_count > 0 && s->xf_num_fmt_ids != NULL) {
        /* We have read-side XF data: rebuild xf_registry from xf_num_fmt_ids */
        for (uint32_t i = 0; i < s->xf_count; i++) {
            uint32_t id = s->xf_num_fmt_ids[i];
            const char *fmt_str = oxl_styles_get_numfmt_str(s, (uint16_t)i);
            if (!fmt_str) {
                fmt_str = "General";
            }
            /* Grow xf_registry if needed */
            if (s->xf_registry_count >= s->xf_registry_cap) {
                uint32_t new_cap = s->xf_registry_cap ? s->xf_registry_cap * 2 : 4;
                void *p = realloc(s->xf_registry, new_cap * sizeof(*s->xf_registry));
                if (!p) return;
                s->xf_registry = p;
                s->xf_registry_cap = new_cap;
            }
            OxlXfRecord *xf = &s->xf_registry[s->xf_registry_count];
            memset(xf, 0, sizeof(*xf));
            xf->fmt_str = strdup(fmt_str);
            xf->num_fmt_id = id;
            s->xf_registry_count++;
        }
        /* Ensure we have at least xf 1 as a date format */
        if (s->xf_registry_count < 2) {
            oxl_styles_get_or_add_xf(s, "YYYY-MM-DD");
        }
        return;
    }

    /* Fresh workbook: xf 0 = General, xf 1 = Date */
    oxl_styles_get_or_add_xf(s, "General");    /* returns 0 */
    oxl_styles_get_or_add_xf(s, "YYYY-MM-DD"); /* returns 1 */
}

/* Helper: emit a color attribute like color rgb="FFFF0000" */
static void emit_color(OxlXmlBuf *b, uint32_t argb) {
    char buf[9];
    snprintf(buf, sizeof(buf), "%08X", argb);
    oxl_xmlbuf_cstr(b, "<color rgb=\"");
    oxl_xmlbuf_cstr(b, buf);
    oxl_xmlbuf_cstr(b, "\"/>");
}

/* Helper: emit a border side element like <left style="thin"><color.../></left> */
static void emit_border_side(OxlXmlBuf *b, const char *tag, const OxlBorderSide *side) {
    oxl_xmlbuf_cstr(b, "<");
    oxl_xmlbuf_cstr(b, tag);
    if (side && side->style) {
        oxl_xmlbuf_cstr(b, " style=\"");
        oxl_xmlbuf_cstr(b, side->style);
        oxl_xmlbuf_cstr(b, "\">");
        if (side->has_color && side->color_rgb) emit_color(b, side->color_rgb);
        oxl_xmlbuf_cstr(b, "</");
        oxl_xmlbuf_cstr(b, tag);
        oxl_xmlbuf_cstr(b, ">");
    } else {
        oxl_xmlbuf_cstr(b, "/>");
    }
}

/* Emit styles.xml content into buffer */
void oxl_write_styles(OxlXmlBuf *b, const OxlStyles *s) {
    oxl_xmlbuf_cstr(b, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

    /* numFmts: only emit custom ones (id >= 164) */
    if (s->custom_fmt_count > 0) {
        oxl_xmlbuf_cstr(b, "<numFmts count=\"");
        oxl_xmlbuf_uint(b, s->custom_fmt_count);
        oxl_xmlbuf_cstr(b, "\">");
        for (uint32_t i = 0; i < s->custom_fmt_count; i++) {
            oxl_xmlbuf_cstr(b, "<numFmt numFmtId=\"");
            oxl_xmlbuf_uint(b, s->custom_fmts[i].id);
            oxl_xmlbuf_cstr(b, "\" formatCode=\"");
            oxl_xmlbuf_attr_val(b, s->custom_fmts[i].fmt_str);
            oxl_xmlbuf_cstr(b, "\"/>");
        }
        oxl_xmlbuf_cstr(b, "</numFmts>");
    }

    /* fonts */
    uint32_t font_count = s->font_count > 0 ? s->font_count : 1;
    oxl_xmlbuf_cstr(b, "<fonts count=\"");
    oxl_xmlbuf_uint(b, font_count);
    oxl_xmlbuf_cstr(b, "\">");
    if (s->font_count == 0) {
        oxl_xmlbuf_cstr(b, "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
    } else {
        for (uint32_t i = 0; i < s->font_count; i++) {
            const OxlFontDef *f = &s->fonts[i];
            oxl_xmlbuf_cstr(b, "<font>");
            if (f->bold)    oxl_xmlbuf_cstr(b, "<b/>");
            if (f->italic)  oxl_xmlbuf_cstr(b, "<i/>");
            if (f->underline == 1) oxl_xmlbuf_cstr(b, "<u/>");
            else if (f->underline == 2) oxl_xmlbuf_cstr(b, "<u val=\"double\"/>");
            oxl_xmlbuf_cstr(b, "<sz val=\"");
            /* emit size as decimal */
            char szb[32];
            double sz = f->size > 0 ? (double)f->size : 11.0;
            snprintf(szb, sizeof(szb), "%.6g", sz);
            oxl_xmlbuf_cstr(b, szb);
            oxl_xmlbuf_cstr(b, "\"/>");
            if (f->color_rgb) emit_color(b, f->color_rgb);
            const char *fname = f->name ? f->name : "Calibri";
            oxl_xmlbuf_cstr(b, "<name val=\"");
            oxl_xmlbuf_attr_val(b, fname);
            oxl_xmlbuf_cstr(b, "\"/>");
            oxl_xmlbuf_cstr(b, "</font>");
        }
    }
    oxl_xmlbuf_cstr(b, "</fonts>");

    /* fills: at least 2 defaults required by OOXML spec */
    uint32_t fill_count = s->fill_count > 2 ? s->fill_count : 2;
    oxl_xmlbuf_cstr(b, "<fills count=\"");
    oxl_xmlbuf_uint(b, fill_count);
    oxl_xmlbuf_cstr(b, "\">");
    if (s->fill_count == 0) {
        oxl_xmlbuf_cstr(b, "<fill><patternFill patternType=\"none\"/></fill>"
                           "<fill><patternFill patternType=\"gray125\"/></fill>");
    } else {
        for (uint32_t i = 0; i < s->fill_count; i++) {
            const OxlFillDef *f = &s->fills[i];
            const char *pt = f->pattern_type ? f->pattern_type : "none";
            oxl_xmlbuf_cstr(b, "<fill><patternFill patternType=\"");
            oxl_xmlbuf_attr_val(b, pt);
            oxl_xmlbuf_cstr(b, "\">");
            if (f->fg_has_color && f->fg_rgb) {
                oxl_xmlbuf_cstr(b, "<fgColor rgb=\"");
                char buf[9]; snprintf(buf, sizeof(buf), "%08X", f->fg_rgb);
                oxl_xmlbuf_cstr(b, buf);
                oxl_xmlbuf_cstr(b, "\"/>");
            }
            if (f->bg_has_color && f->bg_rgb) {
                oxl_xmlbuf_cstr(b, "<bgColor rgb=\"");
                char buf[9]; snprintf(buf, sizeof(buf), "%08X", f->bg_rgb);
                oxl_xmlbuf_cstr(b, buf);
                oxl_xmlbuf_cstr(b, "\"/>");
            }
            oxl_xmlbuf_cstr(b, "</patternFill></fill>");
        }
        /* Pad to 2 if needed */
        for (uint32_t i = s->fill_count; i < 2; i++) {
            if (i == 0) oxl_xmlbuf_cstr(b, "<fill><patternFill patternType=\"none\"/></fill>");
            else        oxl_xmlbuf_cstr(b, "<fill><patternFill patternType=\"gray125\"/></fill>");
        }
    }
    oxl_xmlbuf_cstr(b, "</fills>");

    /* borders: at least 1 default */
    uint32_t border_count = s->border_count > 0 ? s->border_count : 1;
    oxl_xmlbuf_cstr(b, "<borders count=\"");
    oxl_xmlbuf_uint(b, border_count);
    oxl_xmlbuf_cstr(b, "\">");
    if (s->border_count == 0) {
        oxl_xmlbuf_cstr(b, "<border><left/><right/><top/><bottom/><diagonal/></border>");
    } else {
        for (uint32_t i = 0; i < s->border_count; i++) {
            const OxlBorderDef *bd = &s->borders[i];
            oxl_xmlbuf_cstr(b, "<border>");
            emit_border_side(b, "left",     &bd->left);
            emit_border_side(b, "right",    &bd->right);
            emit_border_side(b, "top",      &bd->top);
            emit_border_side(b, "bottom",   &bd->bottom);
            emit_border_side(b, "diagonal", &bd->diagonal);
            oxl_xmlbuf_cstr(b, "</border>");
        }
    }
    oxl_xmlbuf_cstr(b, "</borders>");

    /* cellStyleXfs (master style table, always 1 entry) */
    oxl_xmlbuf_cstr(b, "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

    /* cellXfs: one entry per xf in registry */
    uint32_t count = s->xf_registry_count > 0 ? s->xf_registry_count :
                     (s->xf_count > 0 ? s->xf_count : 2);
    oxl_xmlbuf_cstr(b, "<cellXfs count=\"");
    oxl_xmlbuf_uint(b, count);
    oxl_xmlbuf_cstr(b, "\">");
    for (uint32_t i = 0; i < count; i++) {
        uint32_t fmtid = 0, font_id = 0, fill_id = 0, border_id = 0;
        uint8_t apply_align = 0;
        const OxlAlignDef *align = NULL;

        if (s->xf_registry_count > 0 && i < s->xf_registry_count) {
            const OxlXfRecord *xf = &s->xf_registry[i];
            fmtid     = xf->num_fmt_id;
            font_id   = xf->font_id;
            fill_id   = xf->fill_id;
            border_id = xf->border_id;
            apply_align = xf->apply_alignment;
            if (apply_align) align = &xf->align;
        } else if (s->xf_num_fmt_ids && i < s->xf_count) {
            fmtid = s->xf_num_fmt_ids[i];
        }

        oxl_xmlbuf_cstr(b, "<xf numFmtId=\"");
        oxl_xmlbuf_uint(b, fmtid);
        oxl_xmlbuf_cstr(b, "\" fontId=\"");
        oxl_xmlbuf_uint(b, font_id);
        oxl_xmlbuf_cstr(b, "\" fillId=\"");
        oxl_xmlbuf_uint(b, fill_id);
        oxl_xmlbuf_cstr(b, "\" borderId=\"");
        oxl_xmlbuf_uint(b, border_id);
        oxl_xmlbuf_cstr(b, "\" xfId=\"0\"");

        if (apply_align && align) {
            oxl_xmlbuf_cstr(b, " applyAlignment=\"1\">");
            oxl_xmlbuf_cstr(b, "<alignment");
            if (align->horizontal) {
                oxl_xmlbuf_cstr(b, " horizontal=\"");
                oxl_xmlbuf_cstr(b, align->horizontal);
                oxl_xmlbuf_cstr(b, "\"");
            }
            if (align->vertical) {
                oxl_xmlbuf_cstr(b, " vertical=\"");
                oxl_xmlbuf_cstr(b, align->vertical);
                oxl_xmlbuf_cstr(b, "\"");
            }
            if (align->wrap_text) oxl_xmlbuf_cstr(b, " wrapText=\"1\"");
            if (align->shrink_to_fit) oxl_xmlbuf_cstr(b, " shrinkToFit=\"1\"");
            if (align->indent) {
                char ibuf[32]; snprintf(ibuf, sizeof(ibuf), " indent=\"%d\"", align->indent);
                oxl_xmlbuf_cstr(b, ibuf);
            }
            if (align->text_rotation) {
                char ibuf[32]; snprintf(ibuf, sizeof(ibuf), " textRotation=\"%d\"", align->text_rotation);
                oxl_xmlbuf_cstr(b, ibuf);
            }
            oxl_xmlbuf_cstr(b, "/></xf>");
        } else {
            oxl_xmlbuf_cstr(b, "/>");
        }
    }
    oxl_xmlbuf_cstr(b, "</cellXfs></styleSheet>");
}
