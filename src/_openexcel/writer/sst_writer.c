#include "sst_writer.h"

void oxl_write_sst(OxlXmlBuf *b, const OxlStringTable *sst) {
    oxl_xmlbuf_cstr(b,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        " count=\"");
    oxl_xmlbuf_uint(b, sst->count);
    oxl_xmlbuf_cstr(b, "\" uniqueCount=\"");
    oxl_xmlbuf_uint(b, sst->count);
    oxl_xmlbuf_cstr(b, "\">");

    for (uint32_t i = 0; i < sst->count; i++) {
        oxl_xmlbuf_cstr(b, "<si><t xml:space=\"preserve\">");
        oxl_xmlbuf_text(b, sst->entries[i]);
        oxl_xmlbuf_cstr(b, "</t></si>");
    }
    oxl_xmlbuf_cstr(b, "</sst>");
}
