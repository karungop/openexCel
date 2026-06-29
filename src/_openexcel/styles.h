#pragma once
#include <stdint.h>
#include "writer/xml_gen.h"

typedef struct {
    /* Bit array: bit i = 1 means cell format (xf) index i is a date format */
    uint8_t  *date_xf_bits;
    uint32_t  xf_count;

    /* Phase 2 additions */
    uint32_t *xf_num_fmt_ids;   /* xf_index → numFmtId; length = xf_count */

    /* custom numFmts (id >= 164) — stored on read */
    struct { uint32_t id; char *fmt_str; } *custom_fmts;
    uint32_t   custom_fmt_count;
    uint32_t   custom_fmt_cap;

    /* write-side XF registry: maps format string → xf_index
       Always has at least 2 entries: xf 0 = General, xf 1 = date YYYY-MM-DD */
    struct { char *fmt_str; uint32_t num_fmt_id; } *xf_registry;
    uint32_t   xf_registry_count;   /* == xf_count on write side */
    uint32_t   xf_registry_cap;
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
