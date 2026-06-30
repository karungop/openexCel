#pragma once
#include <stdint.h>
#include "writer/xml_gen.h"

typedef struct {
    char    *name;          /* font family, e.g. "Calibri" */
    float    size;          /* e.g. 11.0 */
    uint32_t color_rgb;     /* ARGB hex value, 0 = auto/theme */
    uint8_t  bold;
    uint8_t  italic;
    uint8_t  underline;     /* 0=none, 1=single, 2=double */
    uint8_t  color_indexed; /* 1 if color_rgb is a theme-based index, unused in basic impl */
} OxlFontDef;

typedef struct {
    char    *pattern_type;  /* "none", "solid", "darkGray", etc. */
    uint32_t fg_rgb;        /* ARGB foreground color */
    uint32_t bg_rgb;        /* ARGB background color */
    uint8_t  fg_has_color;
    uint8_t  bg_has_color;
} OxlFillDef;

typedef struct {
    char    *style;         /* "thin", "medium", "thick", "dashed", "dotted", "double", etc. */
    uint32_t color_rgb;     /* ARGB color */
    uint8_t  has_color;
} OxlBorderSide;

typedef struct {
    OxlBorderSide left;
    OxlBorderSide right;
    OxlBorderSide top;
    OxlBorderSide bottom;
    OxlBorderSide diagonal;
    uint8_t diagonal_up;
    uint8_t diagonal_down;
} OxlBorderDef;

typedef struct {
    char    *horizontal;    /* "left","center","right","general","fill","justify","centerContinuous","distributed" */
    char    *vertical;      /* "top","center","bottom","justify","distributed" */
    int32_t  indent;
    int32_t  text_rotation;
    uint8_t  wrap_text;
    uint8_t  shrink_to_fit;
} OxlAlignDef;

typedef struct {
    char        *fmt_str;       /* keep for backward compat */
    uint32_t     num_fmt_id;
    uint32_t     font_id;       /* index into OxlStyles.fonts[] */
    uint32_t     fill_id;       /* index into OxlStyles.fills[] */
    uint32_t     border_id;     /* index into OxlStyles.borders[] */
    OxlAlignDef  align;
    uint8_t      apply_alignment;
} OxlXfRecord;

typedef struct {
    /* Read-side date detection */
    uint8_t  *date_xf_bits;
    uint32_t  xf_count;

    /* Read-side: xf_index → numFmtId */
    uint32_t *xf_num_fmt_ids;

    /* Custom numFmts (id >= 164) stored on read */
    struct { uint32_t id; char *fmt_str; } *custom_fmts;
    uint32_t   custom_fmt_count;
    uint32_t   custom_fmt_cap;

    /* XF registry: both read-side (populated during parse) and write-side */
    OxlXfRecord *xf_registry;
    uint32_t     xf_registry_count;
    uint32_t     xf_registry_cap;

    /* Font registry */
    OxlFontDef  *fonts;
    uint32_t     font_count;
    uint32_t     font_cap;

    /* Fill registry */
    OxlFillDef  *fills;
    uint32_t     fill_count;
    uint32_t     fill_cap;

    /* Border registry */
    OxlBorderDef *borders;
    uint32_t      border_count;
    uint32_t      border_cap;
} OxlStyles;

void oxl_styles_init(OxlStyles *s);
void oxl_styles_free(OxlStyles *s);

/* Returns 1 if xf_index corresponds to a date/time numFmt. */
int  oxl_styles_is_date(const OxlStyles *s, uint16_t xf_index);

/* Used by xml_styles parser to mark an xf index as date. */
void oxl_styles_set_date(OxlStyles *s, uint32_t xf_index);
void oxl_styles_resize(OxlStyles *s, uint32_t new_xf_count);

/* Returns 1 if a numFmt ID (built-in) is a date format. */
int  oxl_numfmt_id_is_date(uint32_t num_fmt_id);

/* Returns 1 if a custom numFmt format string represents a date. */
int  oxl_numfmt_str_is_date(const char *fmt);

/* Read-side: return format string for an XF index.
   Returns built-in format string for IDs < 164, custom for >= 164.
   Returns NULL for General (numFmtId=0) or unknown. */
const char *oxl_styles_get_numfmt_str(const OxlStyles *s, uint16_t xf_index);

/* Write-side: find or create XF entry for a format string.
   Returns xf_index. Creates new entry if not found.
   fmt_str == NULL or "" → returns 0 (General) */
uint16_t oxl_styles_get_or_add_xf(OxlStyles *s, const char *fmt_str);

/* Called from writer before writing to ensure General and Date XFs exist */
void oxl_styles_init_write_defaults(OxlStyles *s);

/* Emit styles.xml content into buffer (replaces hardcoded STYLES_XML) */
void oxl_write_styles(OxlXmlBuf *b, const OxlStyles *s);

/* Read-side helpers for xml_styles.c */
void oxl_styles_set_xf_numfmt(OxlStyles *s, uint32_t xf_index, uint32_t num_fmt_id);
void oxl_styles_add_custom_fmt(OxlStyles *s, uint32_t id, const char *fmt_str);

/* Accessors for read-side data */
const OxlXfRecord  *oxl_styles_get_xf(const OxlStyles *s, uint16_t xf_index);
const OxlFontDef   *oxl_styles_get_font(const OxlStyles *s, uint32_t font_id);
const OxlFillDef   *oxl_styles_get_fill(const OxlStyles *s, uint32_t fill_id);
const OxlBorderDef *oxl_styles_get_border(const OxlStyles *s, uint32_t border_id);

/* Read-side: add a font/fill/border parsed from XML. Returns the assigned id. */
uint32_t oxl_styles_add_font(OxlStyles *s, const OxlFontDef *font);
uint32_t oxl_styles_add_fill(OxlStyles *s, const OxlFillDef *fill);
uint32_t oxl_styles_add_border(OxlStyles *s, const OxlBorderDef *border);

/* Read-side: after parsing all fonts/fills/borders, set full XF info */
void oxl_styles_set_xf_full(OxlStyles *s, uint32_t xf_index,
                              uint32_t font_id, uint32_t fill_id,
                              uint32_t border_id, const OxlAlignDef *align);

/* Write-side: find-or-create font/fill/border with dedup */
uint32_t oxl_styles_get_or_add_font(OxlStyles *s, const OxlFontDef *font);
uint32_t oxl_styles_get_or_add_fill(OxlStyles *s, const OxlFillDef *fill);
uint32_t oxl_styles_get_or_add_border(OxlStyles *s, const OxlBorderDef *border);

/* Write-side: find-or-create full XF (deduplicates by all fields) */
uint16_t oxl_styles_get_or_add_xf_full(OxlStyles *s, const char *fmt_str,
                                         uint32_t font_id, uint32_t fill_id,
                                         uint32_t border_id, const OxlAlignDef *align);

/* Free heap-owned fields (but not the struct itself) */
void oxl_font_def_free_fields(OxlFontDef *f);
void oxl_fill_def_free_fields(OxlFillDef *f);
void oxl_border_side_free_fields(OxlBorderSide *side);
void oxl_border_def_free_fields(OxlBorderDef *b);
void oxl_align_def_free_fields(OxlAlignDef *a);
