#include "styles.h"
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

/* Built-in Excel numFmt IDs that are date/time formats.
   Reference: ECMA-376 Part 1, §18.8.30 */
static const uint32_t BUILTIN_DATE_FMTS[] = {
    14, 15, 16, 17,          /* m/d/yy, d-mmm-yy, d-mmm, mmm-yy */
    18, 19, 20, 21, 22,      /* h:mm AM/PM, h:mm:ss AM/PM, h:mm, h:mm:ss, m/d/yy h:mm */
    45, 46, 47,              /* mm:ss, [h]:mm:ss, mmss.0 */
};

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
}

void oxl_styles_free(OxlStyles *s) {
    free(s->date_xf_bits);
    s->date_xf_bits = NULL;
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
