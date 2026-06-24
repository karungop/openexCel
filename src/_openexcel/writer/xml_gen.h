#pragma once
#include <stddef.h>

/* Lightweight arena-buffered XML emitter.
   Writes to an in-memory buffer; caller drains with oxl_xmlbuf_take(). */

typedef struct {
    char  *buf;
    size_t len;
    size_t cap;
} OxlXmlBuf;

void oxl_xmlbuf_init(OxlXmlBuf *b);
void oxl_xmlbuf_free(OxlXmlBuf *b);
void oxl_xmlbuf_reset(OxlXmlBuf *b);

/* Low-level append (no escaping). */
void oxl_xmlbuf_raw(OxlXmlBuf *b, const char *s, size_t n);
void oxl_xmlbuf_cstr(OxlXmlBuf *b, const char *s);

/* Append XML-escaped text content (escapes &, <, >, ", '). */
void oxl_xmlbuf_text(OxlXmlBuf *b, const char *s);

/* Append attribute-escaped value (same escaping). */
void oxl_xmlbuf_attr_val(OxlXmlBuf *b, const char *s);

/* Append unsigned integer without snprintf. */
void oxl_xmlbuf_uint(OxlXmlBuf *b, unsigned long v);

/* Append double with up to 15 significant digits. */
void oxl_xmlbuf_double(OxlXmlBuf *b, double v);

/* Append Excel column label: 0→"A", 25→"Z", 26→"AA", etc. */
void oxl_xmlbuf_col_label(OxlXmlBuf *b, unsigned int col);

/* Return the buffer contents (caller must free). Resets the buffer. */
char *oxl_xmlbuf_take(OxlXmlBuf *b, size_t *out_len);
