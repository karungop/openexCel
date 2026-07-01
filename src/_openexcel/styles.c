#include "styles.h"
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <stdio.h>

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

/* ---- Free helpers ---- */

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

/* ---- Init / Free ---- */

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
    s->dxfs = NULL;
    s->dxf_count = 0;
    s->dxf_cap = 0;
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

    if (s->dxfs) {
        for (uint32_t i = 0; i < s->dxf_count; i++) {
            if (s->dxfs[i].font) { oxl_font_def_free_fields(s->dxfs[i].font); free(s->dxfs[i].font); }
            if (s->dxfs[i].fill) { oxl_fill_def_free_fields(s->dxfs[i].fill); free(s->dxfs[i].fill); }
            if (s->dxfs[i].border) { oxl_border_def_free_fields(s->dxfs[i].border); free(s->dxfs[i].border); }
        }
        free(s->dxfs);
        s->dxfs = NULL;
    }
    s->dxf_count = 0;
    s->dxf_cap = 0;

    s->xf_count = 0;
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

/* ---- Font registry ---- */

uint32_t oxl_styles_add_font(OxlStyles *s, const OxlFontDef *font) {
    if (s->font_count >= s->font_cap) {
        uint32_t new_cap = s->font_cap ? s->font_cap * 2 : 4;
        void *p = realloc(s->fonts, new_cap * sizeof(*s->fonts));
        if (!p) return 0;
        s->fonts = p;
        s->font_cap = new_cap;
    }
    OxlFontDef *dst = &s->fonts[s->font_count];
    *dst = *font;
    dst->name = font->name ? strdup(font->name) : NULL;
    return s->font_count++;
}

const OxlFontDef *oxl_styles_get_font(const OxlStyles *s, uint32_t font_id) {
    if (!s->fonts || font_id >= s->font_count) return NULL;
    return &s->fonts[font_id];
}

uint32_t oxl_styles_get_or_add_font(OxlStyles *s, const OxlFontDef *font) {
    for (uint32_t i = 0; i < s->font_count; i++) {
        OxlFontDef *f = &s->fonts[i];
        /* Compare name */
        int name_eq = (f->name == NULL && font->name == NULL) ||
                      (f->name && font->name && strcmp(f->name, font->name) == 0);
        if (!name_eq) continue;
        if (f->size != font->size) continue;
        if (f->bold != font->bold) continue;
        if (f->italic != font->italic) continue;
        if (f->underline != font->underline) continue;
        if (f->color_rgb != font->color_rgb) continue;
        return i;
    }
    return oxl_styles_add_font(s, font);
}

/* ---- Fill registry ---- */

uint32_t oxl_styles_add_fill(OxlStyles *s, const OxlFillDef *fill) {
    if (s->fill_count >= s->fill_cap) {
        uint32_t new_cap = s->fill_cap ? s->fill_cap * 2 : 4;
        void *p = realloc(s->fills, new_cap * sizeof(*s->fills));
        if (!p) return 0;
        s->fills = p;
        s->fill_cap = new_cap;
    }
    OxlFillDef *dst = &s->fills[s->fill_count];
    *dst = *fill;
    dst->pattern_type = fill->pattern_type ? strdup(fill->pattern_type) : NULL;
    return s->fill_count++;
}

const OxlFillDef *oxl_styles_get_fill(const OxlStyles *s, uint32_t fill_id) {
    if (!s->fills || fill_id >= s->fill_count) return NULL;
    return &s->fills[fill_id];
}

uint32_t oxl_styles_get_or_add_fill(OxlStyles *s, const OxlFillDef *fill) {
    for (uint32_t i = 0; i < s->fill_count; i++) {
        OxlFillDef *f = &s->fills[i];
        int pt_eq = (f->pattern_type == NULL && fill->pattern_type == NULL) ||
                    (f->pattern_type && fill->pattern_type &&
                     strcmp(f->pattern_type, fill->pattern_type) == 0);
        if (!pt_eq) continue;
        if (f->fg_rgb != fill->fg_rgb) continue;
        if (f->bg_rgb != fill->bg_rgb) continue;
        if (f->fg_has_color != fill->fg_has_color) continue;
        if (f->bg_has_color != fill->bg_has_color) continue;
        return i;
    }
    return oxl_styles_add_fill(s, fill);
}

/* ---- Border registry ---- */

/* Deep-copy a single border side */
static OxlBorderSide copy_border_side(const OxlBorderSide *src) {
    OxlBorderSide dst = *src;
    dst.style = src->style ? strdup(src->style) : NULL;
    return dst;
}

uint32_t oxl_styles_add_border(OxlStyles *s, const OxlBorderDef *border) {
    if (s->border_count >= s->border_cap) {
        uint32_t new_cap = s->border_cap ? s->border_cap * 2 : 4;
        void *p = realloc(s->borders, new_cap * sizeof(*s->borders));
        if (!p) return 0;
        s->borders = p;
        s->border_cap = new_cap;
    }
    OxlBorderDef *dst = &s->borders[s->border_count];
    dst->left     = copy_border_side(&border->left);
    dst->right    = copy_border_side(&border->right);
    dst->top      = copy_border_side(&border->top);
    dst->bottom   = copy_border_side(&border->bottom);
    dst->diagonal = copy_border_side(&border->diagonal);
    dst->diagonal_up   = border->diagonal_up;
    dst->diagonal_down = border->diagonal_down;
    return s->border_count++;
}

const OxlBorderDef *oxl_styles_get_border(const OxlStyles *s, uint32_t border_id) {
    if (!s->borders || border_id >= s->border_count) return NULL;
    return &s->borders[border_id];
}

static int border_side_eq(const OxlBorderSide *a, const OxlBorderSide *b) {
    int style_eq = (a->style == NULL && b->style == NULL) ||
                   (a->style && b->style && strcmp(a->style, b->style) == 0);
    return style_eq && a->color_rgb == b->color_rgb && a->has_color == b->has_color;
}

uint32_t oxl_styles_get_or_add_border(OxlStyles *s, const OxlBorderDef *border) {
    for (uint32_t i = 0; i < s->border_count; i++) {
        OxlBorderDef *b = &s->borders[i];
        if (!border_side_eq(&b->left,     &border->left))     continue;
        if (!border_side_eq(&b->right,    &border->right))    continue;
        if (!border_side_eq(&b->top,      &border->top))      continue;
        if (!border_side_eq(&b->bottom,   &border->bottom))   continue;
        if (!border_side_eq(&b->diagonal, &border->diagonal)) continue;
        if (b->diagonal_up   != border->diagonal_up)   continue;
        if (b->diagonal_down != border->diagonal_down) continue;
        return i;
    }
    return oxl_styles_add_border(s, border);
}

/* ---- XF registry accessors ---- */

const OxlXfRecord *oxl_styles_get_xf(const OxlStyles *s, uint16_t xf_index) {
    if (!s->xf_registry || xf_index >= s->xf_registry_count) return NULL;
    return &s->xf_registry[xf_index];
}

/* Read-side: after parsing all fonts/fills/borders, set full XF info */
void oxl_styles_set_xf_full(OxlStyles *s, uint32_t xf_index,
                              uint32_t font_id, uint32_t fill_id,
                              uint32_t border_id, const OxlAlignDef *align) {
    /* Grow xf_registry if needed */
    if (xf_index >= s->xf_registry_cap) {
        uint32_t new_cap = s->xf_registry_cap ? s->xf_registry_cap * 2 : 4;
        while (new_cap <= xf_index) new_cap *= 2;
        void *p = realloc(s->xf_registry, new_cap * sizeof(*s->xf_registry));
        if (!p) return;
        /* Zero-fill new entries */
        memset((char *)p + s->xf_registry_cap * sizeof(*s->xf_registry), 0,
               (new_cap - s->xf_registry_cap) * sizeof(*s->xf_registry));
        s->xf_registry = p;
        s->xf_registry_cap = new_cap;
    }
    if (xf_index >= s->xf_registry_count) {
        s->xf_registry_count = xf_index + 1;
    }

    OxlXfRecord *rec = &s->xf_registry[xf_index];
    rec->font_id   = font_id;
    rec->fill_id   = fill_id;
    rec->border_id = border_id;

    if (align) {
        /* Free old strings if any */
        free(rec->align.horizontal);
        free(rec->align.vertical);
        rec->align = *align;
        rec->align.horizontal = align->horizontal ? strdup(align->horizontal) : NULL;
        rec->align.vertical   = align->vertical   ? strdup(align->vertical)   : NULL;
        rec->apply_alignment = (align->horizontal || align->vertical ||
                                align->wrap_text || align->shrink_to_fit ||
                                align->indent || align->text_rotation) ? 1 : 0;
    }
}

/* ---- Write-side: find or create an XF entry for a format string ---- */

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
    /* Zero-init the new record, then set fields */
    memset(&s->xf_registry[new_idx], 0, sizeof(s->xf_registry[new_idx]));
    s->xf_registry[new_idx].fmt_str = strdup(fmt_str);
    s->xf_registry[new_idx].num_fmt_id = num_fmt_id;
    /* font_id, fill_id, border_id default to 0 */
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

/* Write-side: find-or-create full XF with all style fields */
uint16_t oxl_styles_get_or_add_xf_full(OxlStyles *s, const char *fmt_str,
                                         uint32_t font_id, uint32_t fill_id,
                                         uint32_t border_id, const OxlAlignDef *align) {
    /* Normalize fmt_str */
    const char *fs = (fmt_str && fmt_str[0]) ? fmt_str : "General";

    /* Search existing registry for full match */
    for (uint32_t i = 0; i < s->xf_registry_count; i++) {
        OxlXfRecord *r = &s->xf_registry[i];
        if (!r->fmt_str || strcmp(r->fmt_str, fs) != 0) continue;
        if (r->font_id   != font_id)   continue;
        if (r->fill_id   != fill_id)   continue;
        if (r->border_id != border_id) continue;
        /* Compare alignment */
        if (align) {
            int h_eq = (r->align.horizontal == NULL && align->horizontal == NULL) ||
                       (r->align.horizontal && align->horizontal &&
                        strcmp(r->align.horizontal, align->horizontal) == 0);
            int v_eq = (r->align.vertical == NULL && align->vertical == NULL) ||
                       (r->align.vertical && align->vertical &&
                        strcmp(r->align.vertical, align->vertical) == 0);
            if (!h_eq) continue;
            if (!v_eq) continue;
            if (r->align.wrap_text     != align->wrap_text)     continue;
            if (r->align.indent        != align->indent)        continue;
            if (r->align.text_rotation != align->text_rotation) continue;
            if (r->align.shrink_to_fit != align->shrink_to_fit) continue;
        } else {
            /* align == NULL means default (all zeros): check record has no alignment */
            if (r->align.horizontal || r->align.vertical ||
                r->align.wrap_text || r->align.indent ||
                r->align.text_rotation || r->align.shrink_to_fit) continue;
        }
        return (uint16_t)i;
    }

    /* Determine numFmtId */
    uint32_t num_fmt_id = 0;
    int is_general = (strcmp(fs, "General") == 0);

    if (!is_general) {
        int found_builtin = 0;
        for (size_t i = 0; i < BUILTIN_FMTS_COUNT; i++) {
            if (strcasecmp(BUILTIN_FMTS[i].str, fs) == 0) {
                num_fmt_id = BUILTIN_FMTS[i].id;
                found_builtin = 1;
                break;
            }
        }
        if (!found_builtin) {
            /* Check if already in custom_fmts */
            int found_custom = 0;
            for (uint32_t i = 0; i < s->custom_fmt_count; i++) {
                if (strcmp(s->custom_fmts[i].fmt_str, fs) == 0) {
                    num_fmt_id = s->custom_fmts[i].id;
                    found_custom = 1;
                    break;
                }
            }
            if (!found_custom) {
                num_fmt_id = 164 + s->custom_fmt_count;
                oxl_styles_add_custom_fmt(s, num_fmt_id, fs);
            }
        }
    }

    /* Grow xf_registry if needed */
    if (s->xf_registry_count >= s->xf_registry_cap) {
        uint32_t new_cap = s->xf_registry_cap ? s->xf_registry_cap * 2 : 4;
        void *p = realloc(s->xf_registry, new_cap * sizeof(*s->xf_registry));
        if (!p) return 0;
        s->xf_registry = p;
        s->xf_registry_cap = new_cap;
    }

    uint32_t new_idx = s->xf_registry_count;
    OxlXfRecord *rec = &s->xf_registry[new_idx];
    memset(rec, 0, sizeof(*rec));
    rec->fmt_str   = strdup(fs);
    rec->num_fmt_id = num_fmt_id;
    rec->font_id   = font_id;
    rec->fill_id   = fill_id;
    rec->border_id = border_id;

    if (align) {
        rec->align = *align;
        rec->align.horizontal = align->horizontal ? strdup(align->horizontal) : NULL;
        rec->align.vertical   = align->vertical   ? strdup(align->vertical)   : NULL;
        rec->apply_alignment  = (align->horizontal || align->vertical ||
                                 align->wrap_text || align->shrink_to_fit ||
                                 align->indent || align->text_rotation) ? 1 : 0;
    }
    s->xf_registry_count++;

    /* Keep xf_count, date_xf_bits, and xf_num_fmt_ids in sync */
    oxl_styles_resize(s, new_idx + 1);
    if (s->xf_num_fmt_ids) s->xf_num_fmt_ids[new_idx] = num_fmt_id;

    if (!is_general && oxl_numfmt_str_is_date(fs)) {
        oxl_styles_set_date(s, new_idx);
    }

    return (uint16_t)new_idx;
}

/* Initialize write-side defaults (xf 0 = General, xf 1 = date YYYY-MM-DD) */
void oxl_styles_init_write_defaults(OxlStyles *s) {
    if (s->xf_registry_count > 0) return;

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
            OxlXfRecord *rec = &s->xf_registry[s->xf_registry_count];
            memset(rec, 0, sizeof(*rec));
            rec->fmt_str = strdup(fmt_str);
            rec->num_fmt_id = id;
            s->xf_registry_count++;
        }
        /* Ensure we have at least xf 1 as a date format */
        if (s->xf_registry_count < 2) {
            oxl_styles_get_or_add_xf(s, "YYYY-MM-DD");
        }
        goto setup_defaults;
    }

    /* Fresh workbook: xf 0 = General, xf 1 = Date */
    oxl_styles_get_or_add_xf(s, "General");    /* returns 0 */
    oxl_styles_get_or_add_xf(s, "YYYY-MM-DD"); /* returns 1 */

setup_defaults:
    /* Ensure default font (Calibri 11) exists at index 0 */
    if (s->font_count == 0) {
        OxlFontDef default_font = {0};
        default_font.name = strdup("Calibri");
        default_font.size = 11.0f;
        oxl_styles_add_font(s, &default_font);
        free(default_font.name);
    }

    /* Ensure two required fills exist (spec requires at least "none" and "gray125") */
    if (s->fill_count == 0) {
        OxlFillDef f0 = {0}; f0.pattern_type = strdup("none");
        oxl_styles_add_fill(s, &f0); free(f0.pattern_type);
        OxlFillDef f1 = {0}; f1.pattern_type = strdup("gray125");
        oxl_styles_add_fill(s, &f1); free(f1.pattern_type);
    }

    /* Ensure default border exists at index 0 */
    if (s->border_count == 0) {
        OxlBorderDef default_border = {0};
        oxl_styles_add_border(s, &default_border);
    }
}

/* ---- XML writer helpers ---- */

static void write_argb_color(OxlXmlBuf *b, uint32_t argb) {
    char hex[9];
    snprintf(hex, sizeof(hex), "%08X", argb);
    oxl_xmlbuf_cstr(b, hex);
}

static void write_border_side(OxlXmlBuf *b, const char *tag, const OxlBorderSide *side) {
    if (side->style && side->style[0]) {
        oxl_xmlbuf_cstr(b, "<");
        oxl_xmlbuf_cstr(b, tag);
        oxl_xmlbuf_cstr(b, " style=\"");
        oxl_xmlbuf_cstr(b, side->style);
        oxl_xmlbuf_cstr(b, "\">");
        if (side->has_color) {
            oxl_xmlbuf_cstr(b, "<color rgb=\"");
            write_argb_color(b, side->color_rgb);
            oxl_xmlbuf_cstr(b, "\"/>");
        }
        oxl_xmlbuf_cstr(b, "</");
        oxl_xmlbuf_cstr(b, tag);
        oxl_xmlbuf_cstr(b, ">");
    } else {
        oxl_xmlbuf_cstr(b, "<");
        oxl_xmlbuf_cstr(b, tag);
        oxl_xmlbuf_cstr(b, "/>");
    }
}

/* Phase 16: DXF helpers */

uint32_t oxl_styles_add_dxf(OxlStyles *s, const OxlFontDef *font,
                              const OxlFillDef *fill, const OxlBorderDef *border) {
    if (s->dxf_count >= s->dxf_cap) {
        uint32_t cap = s->dxf_cap ? s->dxf_cap * 2 : 8;
        OxlDxf *p = realloc(s->dxfs, cap * sizeof(OxlDxf));
        if (!p) return 0;
        s->dxfs = p;
        s->dxf_cap = cap;
    }
    OxlDxf *d = &s->dxfs[s->dxf_count];
    memset(d, 0, sizeof(*d));
    if (font) {
        d->font = calloc(1, sizeof(OxlFontDef));
        if (d->font) {
            *d->font = *font;
            d->font->name = font->name ? strdup(font->name) : NULL;
        }
    }
    if (fill) {
        d->fill = calloc(1, sizeof(OxlFillDef));
        if (d->fill) {
            *d->fill = *fill;
            d->fill->pattern_type = fill->pattern_type ? strdup(fill->pattern_type) : NULL;
        }
    }
    if (border) {
        d->border = calloc(1, sizeof(OxlBorderDef));
        if (d->border) {
            *d->border = *border;
            d->border->left.style     = border->left.style     ? strdup(border->left.style)     : NULL;
            d->border->right.style    = border->right.style    ? strdup(border->right.style)    : NULL;
            d->border->top.style      = border->top.style      ? strdup(border->top.style)      : NULL;
            d->border->bottom.style   = border->bottom.style   ? strdup(border->bottom.style)   : NULL;
            d->border->diagonal.style = border->diagonal.style ? strdup(border->diagonal.style) : NULL;
        }
    }
    return s->dxf_count++;
}

const OxlDxf *oxl_styles_get_dxf(const OxlStyles *s, uint32_t dxf_id) {
    if (!s->dxfs || dxf_id >= s->dxf_count) return NULL;
    return &s->dxfs[dxf_id];
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
    if (s->font_count > 0) {
        oxl_xmlbuf_cstr(b, "<fonts count=\"");
        oxl_xmlbuf_uint(b, s->font_count);
        oxl_xmlbuf_cstr(b, "\">");
        for (uint32_t i = 0; i < s->font_count; i++) {
            const OxlFontDef *f = &s->fonts[i];
            oxl_xmlbuf_cstr(b, "<font>");
            if (f->size > 0) {
                /* Emit size: use integer if it's a whole number, else float */
                oxl_xmlbuf_cstr(b, "<sz val=\"");
                char szstr[32];
                if (f->size == (float)(int)f->size)
                    snprintf(szstr, sizeof(szstr), "%d", (int)f->size);
                else
                    snprintf(szstr, sizeof(szstr), "%.2f", f->size);
                oxl_xmlbuf_cstr(b, szstr);
                oxl_xmlbuf_cstr(b, "\"/>");
            }
            if (f->name) {
                oxl_xmlbuf_cstr(b, "<name val=\"");
                oxl_xmlbuf_attr_val(b, f->name);
                oxl_xmlbuf_cstr(b, "\"/>");
            }
            if (f->bold)    oxl_xmlbuf_cstr(b, "<b/>");
            if (f->italic)  oxl_xmlbuf_cstr(b, "<i/>");
            if (f->underline == 1) oxl_xmlbuf_cstr(b, "<u val=\"single\"/>");
            else if (f->underline == 2) oxl_xmlbuf_cstr(b, "<u val=\"double\"/>");
            if (f->color_rgb != 0) {
                oxl_xmlbuf_cstr(b, "<color rgb=\"");
                write_argb_color(b, f->color_rgb);
                oxl_xmlbuf_cstr(b, "\"/>");
            }
            oxl_xmlbuf_cstr(b, "</font>");
        }
        oxl_xmlbuf_cstr(b, "</fonts>");
    } else {
        /* Fallback: single default font */
        oxl_xmlbuf_cstr(b, "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>");
    }

    /* fills */
    if (s->fill_count > 0) {
        oxl_xmlbuf_cstr(b, "<fills count=\"");
        oxl_xmlbuf_uint(b, s->fill_count);
        oxl_xmlbuf_cstr(b, "\">");
        for (uint32_t i = 0; i < s->fill_count; i++) {
            const OxlFillDef *f = &s->fills[i];
            oxl_xmlbuf_cstr(b, "<fill>");
            if (f->pattern_type && f->pattern_type[0]) {
                oxl_xmlbuf_cstr(b, "<patternFill patternType=\"");
                oxl_xmlbuf_cstr(b, f->pattern_type);
                oxl_xmlbuf_cstr(b, "\">");
                if (f->fg_has_color) {
                    oxl_xmlbuf_cstr(b, "<fgColor rgb=\"");
                    write_argb_color(b, f->fg_rgb);
                    oxl_xmlbuf_cstr(b, "\"/>");
                }
                if (f->fg_has_color || (f->pattern_type && strcmp(f->pattern_type, "solid") == 0)) {
                    /* Excel requires bgColor indexed="64" for solid fills */
                    oxl_xmlbuf_cstr(b, "<bgColor indexed=\"64\"/>");
                }
                oxl_xmlbuf_cstr(b, "</patternFill>");
            } else {
                oxl_xmlbuf_cstr(b, "<patternFill patternType=\"none\"/>");
            }
            oxl_xmlbuf_cstr(b, "</fill>");
        }
        oxl_xmlbuf_cstr(b, "</fills>");
    } else {
        /* Fallback: 2 default fills */
        oxl_xmlbuf_cstr(b, "<fills count=\"2\">"
            "<fill><patternFill patternType=\"none\"/></fill>"
            "<fill><patternFill patternType=\"gray125\"/></fill>"
            "</fills>");
    }

    /* borders */
    if (s->border_count > 0) {
        oxl_xmlbuf_cstr(b, "<borders count=\"");
        oxl_xmlbuf_uint(b, s->border_count);
        oxl_xmlbuf_cstr(b, "\">");
        for (uint32_t i = 0; i < s->border_count; i++) {
            const OxlBorderDef *bd = &s->borders[i];
            oxl_xmlbuf_cstr(b, "<border>");
            write_border_side(b, "left",     &bd->left);
            write_border_side(b, "right",    &bd->right);
            write_border_side(b, "top",      &bd->top);
            write_border_side(b, "bottom",   &bd->bottom);
            write_border_side(b, "diagonal", &bd->diagonal);
            oxl_xmlbuf_cstr(b, "</border>");
        }
        oxl_xmlbuf_cstr(b, "</borders>");
    } else {
        /* Fallback: 1 default border */
        oxl_xmlbuf_cstr(b, "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");
    }

    /* cellStyleXfs (master style table, always 1 entry) */
    oxl_xmlbuf_cstr(b, "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

    /* cellXfs: one entry per xf in registry */
    uint32_t count = s->xf_registry_count > 0 ? s->xf_registry_count :
                     (s->xf_count > 0 ? s->xf_count : 2);
    oxl_xmlbuf_cstr(b, "<cellXfs count=\"");
    oxl_xmlbuf_uint(b, count);
    oxl_xmlbuf_cstr(b, "\">");
    for (uint32_t i = 0; i < count; i++) {
        uint32_t fmtid    = 0;
        uint32_t font_id  = 0;
        uint32_t fill_id  = 0;
        uint32_t border_id = 0;
        uint8_t  apply_align = 0;
        const OxlAlignDef *align = NULL;

        if (s->xf_registry_count > 0 && i < s->xf_registry_count) {
            const OxlXfRecord *r = &s->xf_registry[i];
            fmtid     = r->num_fmt_id;
            font_id   = r->font_id;
            fill_id   = r->fill_id;
            border_id = r->border_id;
            apply_align = r->apply_alignment;
            if (apply_align) align = &r->align;
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
            if (align->wrap_text) {
                oxl_xmlbuf_cstr(b, " wrapText=\"1\"");
            }
            if (align->indent) {
                oxl_xmlbuf_cstr(b, " indent=\"");
                oxl_xmlbuf_uint(b, (uint32_t)align->indent);
                oxl_xmlbuf_cstr(b, "\"");
            }
            if (align->text_rotation) {
                oxl_xmlbuf_cstr(b, " textRotation=\"");
                oxl_xmlbuf_uint(b, (uint32_t)align->text_rotation);
                oxl_xmlbuf_cstr(b, "\"");
            }
            if (align->shrink_to_fit) {
                oxl_xmlbuf_cstr(b, " shrinkToFit=\"1\"");
            }
            oxl_xmlbuf_cstr(b, "/></xf>");
        } else {
            oxl_xmlbuf_cstr(b, "/>");
        }
    }
    oxl_xmlbuf_cstr(b, "</cellXfs>");

    /* Phase 16: <dxfs> — differential formatting records for conditional formatting */
    if (s->dxf_count > 0) {
        oxl_xmlbuf_cstr(b, "<dxfs count=\"");
        oxl_xmlbuf_uint(b, s->dxf_count);
        oxl_xmlbuf_cstr(b, "\">");
        for (uint32_t i = 0; i < s->dxf_count; i++) {
            const OxlDxf *dxf = &s->dxfs[i];
            oxl_xmlbuf_cstr(b, "<dxf>");
            if (dxf->font) {
                const OxlFontDef *f = dxf->font;
                oxl_xmlbuf_cstr(b, "<font>");
                if (f->bold)   oxl_xmlbuf_cstr(b, "<b/>");
                if (f->italic) oxl_xmlbuf_cstr(b, "<i/>");
                if (f->underline == 1) oxl_xmlbuf_cstr(b, "<u/>");
                else if (f->underline == 2) oxl_xmlbuf_cstr(b, "<u val=\"double\"/>");
                if (f->size > 0) {
                    char buf[32]; snprintf(buf, sizeof(buf), "%.6g", (double)f->size);
                    oxl_xmlbuf_cstr(b, "<sz val=\""); oxl_xmlbuf_cstr(b, buf); oxl_xmlbuf_cstr(b, "\"/>");
                }
                if (f->color_rgb) {
                    oxl_xmlbuf_cstr(b, "<color rgb=\""); write_argb_color(b, f->color_rgb); oxl_xmlbuf_cstr(b, "\"/>");
                }
                if (f->name) {
                    oxl_xmlbuf_cstr(b, "<name val=\""); oxl_xmlbuf_attr_val(b, f->name); oxl_xmlbuf_cstr(b, "\"/>");
                }
                oxl_xmlbuf_cstr(b, "</font>");
            }
            if (dxf->fill) {
                const OxlFillDef *f = dxf->fill;
                oxl_xmlbuf_cstr(b, "<fill><patternFill");
                if (f->pattern_type) {
                    oxl_xmlbuf_cstr(b, " patternType=\"");
                    oxl_xmlbuf_cstr(b, f->pattern_type);
                    oxl_xmlbuf_cstr(b, "\"");
                }
                oxl_xmlbuf_raw(b, ">", 1);
                if (f->fg_has_color && f->fg_rgb) {
                    oxl_xmlbuf_cstr(b, "<fgColor rgb=\""); write_argb_color(b, f->fg_rgb); oxl_xmlbuf_cstr(b, "\"/>");
                }
                if (f->bg_has_color && f->bg_rgb) {
                    oxl_xmlbuf_cstr(b, "<bgColor rgb=\""); write_argb_color(b, f->bg_rgb); oxl_xmlbuf_cstr(b, "\"/>");
                }
                oxl_xmlbuf_cstr(b, "</patternFill></fill>");
            }
            if (dxf->border) {
                const OxlBorderDef *bd = dxf->border;
                oxl_xmlbuf_cstr(b, "<border>");
                write_border_side(b, "left",   &bd->left);
                write_border_side(b, "right",  &bd->right);
                write_border_side(b, "top",    &bd->top);
                write_border_side(b, "bottom", &bd->bottom);
                oxl_xmlbuf_cstr(b, "</border>");
            }
            oxl_xmlbuf_cstr(b, "</dxf>");
        }
        oxl_xmlbuf_cstr(b, "</dxfs>");
    }

    oxl_xmlbuf_cstr(b, "</styleSheet>");
}
