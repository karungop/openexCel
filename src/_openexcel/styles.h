#pragma once
#include <stdint.h>

typedef struct {
    /* Bit array: bit i = 1 means cell format (xf) index i is a date format */
    uint8_t  *date_xf_bits;
    uint32_t  xf_count;
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
