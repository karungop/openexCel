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
        for (uint32_t i = 0; i < s->xf_registry_count; i++)
            free(s->xf_registry[i].fmt_str);
        free(s->xf_registry);
        s->xf_registry = NULL;
    }
    s->xf_registry_count = 0;
    s->xf_registry_cap = 0;

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
            s->xf_registry[s->xf_registry_count].fmt_str = strdup(fmt_str);
            s->xf_registry[s->xf_registry_count].num_fmt_id = id;
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

    /* fonts (always 1 default) */
    oxl_xmlbuf_cstr(b, "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>");

    /* fills (always 2 defaults required by spec) */
    oxl_xmlbuf_cstr(b, "<fills count=\"2\">"
        "<fill><patternFill patternType=\"none\"/></fill>"
        "<fill><patternFill patternType=\"gray125\"/></fill>"
        "</fills>");

    /* borders (always 1 default) */
    oxl_xmlbuf_cstr(b, "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");

    /* cellStyleXfs (master style table, always 1 entry) */
    oxl_xmlbuf_cstr(b, "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

    /* cellXfs: one entry per xf in registry */
    uint32_t count = s->xf_registry_count > 0 ? s->xf_registry_count :
                     (s->xf_count > 0 ? s->xf_count : 2);
    oxl_xmlbuf_cstr(b, "<cellXfs count=\"");
    oxl_xmlbuf_uint(b, count);
    oxl_xmlbuf_cstr(b, "\">");
    for (uint32_t i = 0; i < count; i++) {
        uint32_t fmtid = 0;
        if (s->xf_registry_count > 0 && i < s->xf_registry_count) {
            fmtid = s->xf_registry[i].num_fmt_id;
        } else if (s->xf_num_fmt_ids && i < s->xf_count) {
            fmtid = s->xf_num_fmt_ids[i];
        }
        oxl_xmlbuf_cstr(b, "<xf numFmtId=\"");
        oxl_xmlbuf_uint(b, fmtid);
        oxl_xmlbuf_cstr(b, "\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
    }
    oxl_xmlbuf_cstr(b, "</cellXfs></styleSheet>");
}
