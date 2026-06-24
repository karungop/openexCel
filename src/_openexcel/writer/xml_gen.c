#include "xml_gen.h"
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <math.h>

void oxl_xmlbuf_init(OxlXmlBuf *b) {
    b->buf = NULL;
    b->len = 0;
    b->cap = 0;
}

void oxl_xmlbuf_free(OxlXmlBuf *b) {
    free(b->buf);
    b->buf = NULL;
    b->len = 0;
    b->cap = 0;
}

void oxl_xmlbuf_reset(OxlXmlBuf *b) {
    b->len = 0;
}

static void ensure(OxlXmlBuf *b, size_t need) {
    if (b->len + need <= b->cap) return;
    size_t cap = b->cap ? b->cap * 2 : 8192;
    while (cap < b->len + need) cap *= 2;
    char *p = realloc(b->buf, cap);
    if (!p) return; /* OOM: just drop data */
    b->buf = p;
    b->cap = cap;
}

void oxl_xmlbuf_raw(OxlXmlBuf *b, const char *s, size_t n) {
    ensure(b, n);
    memcpy(b->buf + b->len, s, n);
    b->len += n;
}

void oxl_xmlbuf_cstr(OxlXmlBuf *b, const char *s) {
    oxl_xmlbuf_raw(b, s, strlen(s));
}

void oxl_xmlbuf_text(OxlXmlBuf *b, const char *s) {
    for (; *s; s++) {
        switch (*s) {
        case '&':  oxl_xmlbuf_raw(b, "&amp;",  5); break;
        case '<':  oxl_xmlbuf_raw(b, "&lt;",   4); break;
        case '>':  oxl_xmlbuf_raw(b, "&gt;",   4); break;
        case '"':  oxl_xmlbuf_raw(b, "&quot;", 6); break;
        default:   oxl_xmlbuf_raw(b, s, 1); break;
        }
    }
}

void oxl_xmlbuf_attr_val(OxlXmlBuf *b, const char *s) {
    oxl_xmlbuf_text(b, s); /* same escaping for attribute values */
}

void oxl_xmlbuf_uint(OxlXmlBuf *b, unsigned long v) {
    char tmp[24];
    int  n = 0;
    if (v == 0) { oxl_xmlbuf_raw(b, "0", 1); return; }
    while (v) { tmp[n++] = (char)('0' + v % 10); v /= 10; }
    /* reverse */
    for (int i = 0, j = n-1; i < j; i++, j--) {
        char t = tmp[i]; tmp[i] = tmp[j]; tmp[j] = t;
    }
    oxl_xmlbuf_raw(b, tmp, (size_t)n);
}

void oxl_xmlbuf_double(OxlXmlBuf *b, double v) {
    char tmp[32];
    int n = snprintf(tmp, sizeof(tmp), "%.15g", v);
    if (n > 0) oxl_xmlbuf_raw(b, tmp, (size_t)n);
}

void oxl_xmlbuf_col_label(OxlXmlBuf *b, unsigned int col) {
    char tmp[8];
    int  n = 0;
    col++; /* 0-based → 1-based */
    while (col > 0) {
        col--;
        tmp[n++] = (char)('A' + col % 26);
        col /= 26;
    }
    /* reverse */
    for (int i = 0, j = n-1; i < j; i++, j--) {
        char t = tmp[i]; tmp[i] = tmp[j]; tmp[j] = t;
    }
    oxl_xmlbuf_raw(b, tmp, (size_t)n);
}

char *oxl_xmlbuf_take(OxlXmlBuf *b, size_t *out_len) {
    char *p = b->buf;
    *out_len = b->len;
    b->buf = NULL;
    b->len = 0;
    b->cap = 0;
    return p;
}
