#include "writer.h"
#include "zip_writer.h"
#include "xml_gen.h"
#include "sst_writer.h"
#include "sheet_writer.h"
#include "../cell.h"
#include "../styles.h"
#include "../worksheet.h"
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

#define ERR(fmt, ...) \
    do { if (err_buf) snprintf(err_buf, err_cap, fmt, ##__VA_ARGS__); } while(0)

/* ---- Static XML templates ---- */

static const char CONTENT_TYPES_TMPL[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
    "<Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
    "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
    "<Override PartName=\"/xl/sharedStrings.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
    "<Override PartName=\"/xl/styles.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
    "%s"   /* per-sheet overrides injected here */
    "</Types>";

static const char RELS_TMPL[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>"
    "</Relationships>";

/* styles.xml is now generated dynamically by oxl_write_styles() */

static int sheet_has_hyperlinks(const OxlWorksheet *ws) {
    for (uint32_t i = 0; i < ws->cell_count; i++)
        if (ws->cells[i].hyperlink) return 1;
    return 0;
}

static void write_sheet_rels_buf(OxlXmlBuf *b, const OxlWorksheet *ws) {
    oxl_xmlbuf_cstr(b,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    uint32_t rid = 1;
    for (uint32_t i = 0; i < ws->cell_count; i++) {
        if (!ws->cells[i].hyperlink) continue;
        oxl_xmlbuf_cstr(b, "<Relationship Id=\"rId");
        oxl_xmlbuf_uint(b, rid++);
        oxl_xmlbuf_cstr(b,
            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/hyperlink\" Target=\"");
        oxl_xmlbuf_attr_val(b, ws->cells[i].hyperlink);
        oxl_xmlbuf_cstr(b, "\" TargetMode=\"External\"/>");
    }
    oxl_xmlbuf_cstr(b, "</Relationships>");
}

static void intern_worksheet_strings(OxlWorkbook *wb, OxlWorksheet *ws) {
    for (uint32_t i = 0; i < ws->cell_count; i++) {
        OxlCell *c = &ws->cells[i];
        if (c->type == OXL_CELL_INLINE_STR && c->v.s_inline) {
            uint32_t idx = oxl_sst_intern(&wb->sst, c->v.s_inline);
            free(c->v.s_inline);
            c->v.s_inline = NULL;
            c->type  = OXL_CELL_STRING;
            c->v.s_idx = idx;
        }
    }
}

int oxl_write_workbook(OxlWorkbook *wb, const char *path,
                       char *err_buf, size_t err_cap) {
    OxlZipWriter *zw = oxl_zipw_open(path);
    if (!zw) { ERR("Cannot create file: %s", path); return -1; }

    OxlXmlBuf b;
    oxl_xmlbuf_init(&b);
    int rc = 0;

#define ADD(name) \
    do { size_t _len; char *_p = oxl_xmlbuf_take(&b, &_len); \
         if (oxl_zipw_add(zw, name, _p, _len) != 0) { free(_p); ERR("ZIP write failed: %s", name); rc = -1; goto done; } \
         free(_p); } while(0)

    /* 1. Intern all inline strings across all sheets */
    for (uint32_t i = 0; i < wb->sheet_count; i++)
        intern_worksheet_strings(wb, wb->sheets[i]);

    /* 2. [Content_Types].xml — build per-sheet overrides first */
    {
        OxlXmlBuf sheet_ct;
        oxl_xmlbuf_init(&sheet_ct);
        for (uint32_t i = 0; i < wb->sheet_count; i++) {
            oxl_xmlbuf_cstr(&sheet_ct, "<Override PartName=\"/xl/worksheets/sheet");
            oxl_xmlbuf_uint(&sheet_ct, i + 1);
            oxl_xmlbuf_cstr(&sheet_ct,
                ".xml\" ContentType=\"application/"
                "vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        }
        size_t ct_len; char *ct = oxl_xmlbuf_take(&sheet_ct, &ct_len);
        /* We need null-terminated ct for snprintf */
        char *ct_z = realloc(ct, ct_len + 1); if (ct_z) { ct_z[ct_len] = '\0'; ct = ct_z; }
        int need = snprintf(NULL, 0, CONTENT_TYPES_TMPL, ct) + 1;
        char *full = malloc((size_t)need);
        if (full) { snprintf(full, (size_t)need, CONTENT_TYPES_TMPL, ct); }
        free(ct);
        if (!full) { rc = -1; goto done; }
        oxl_xmlbuf_raw(&b, full, (size_t)(need - 1));
        free(full);
    }
    ADD("[Content_Types].xml");

    /* 3. _rels/.rels */
    oxl_xmlbuf_cstr(&b, RELS_TMPL);
    ADD("_rels/.rels");

    /* 4. xl/_rels/workbook.xml.rels */
    oxl_xmlbuf_cstr(&b,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    oxl_xmlbuf_cstr(&b,
        "<Relationship Id=\"rId0\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" "
        "Target=\"sharedStrings.xml\"/>");
    oxl_xmlbuf_cstr(&b,
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
        "Target=\"styles.xml\"/>");
    for (uint32_t i = 0; i < wb->sheet_count; i++) {
        oxl_xmlbuf_cstr(&b, "<Relationship Id=\"rId");
        oxl_xmlbuf_uint(&b, i + 2);
        oxl_xmlbuf_cstr(&b,
            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
            "Target=\"worksheets/sheet");
        oxl_xmlbuf_uint(&b, i + 1);
        oxl_xmlbuf_cstr(&b, ".xml\"/>");
    }
    oxl_xmlbuf_cstr(&b, "</Relationships>");
    ADD("xl/_rels/workbook.xml.rels");

    /* 5. xl/sharedStrings.xml */
    oxl_write_sst(&b, &wb->sst);
    ADD("xl/sharedStrings.xml");

    /* Phase 16: assign dxf_ids for all CF rules that have inline styling */
    for (uint32_t si = 0; si < wb->sheet_count; si++) {
        OxlWorksheet *ws2 = wb->sheets[si];
        for (uint32_t ci = 0; ci < ws2->cf_count; ci++) {
            OxlCf *cf = &ws2->cond_fmts[ci];
            for (uint32_t ri = 0; ri < cf->rule_count; ri++) {
                OxlCfRule *rule = &cf->rules[ri];
                if (rule->dxf_id >= 0) continue;  /* already assigned (from read) */
                if (rule->font || rule->fill || rule->border) {
                    rule->dxf_id = (int32_t)oxl_styles_add_dxf(
                        &wb->styles, rule->font, rule->fill, rule->border);
                }
            }
        }
    }

    /* 6. xl/styles.xml */
    oxl_styles_init_write_defaults(&wb->styles);
    oxl_write_styles(&b, &wb->styles);
    ADD("xl/styles.xml");

    /* 7. xl/worksheets/sheetN.xml */
    for (uint32_t i = 0; i < wb->sheet_count; i++) {
        char entry[64];
        snprintf(entry, sizeof(entry), "xl/worksheets/sheet%u.xml", i + 1);
        oxl_write_sheet(&b, wb->sheets[i], wb, 1);
        ADD(entry);

        /* Write sheet rels file if this sheet has hyperlinks */
        if (sheet_has_hyperlinks(wb->sheets[i])) {
            char rels_entry[128];
            snprintf(rels_entry, sizeof(rels_entry),
                     "xl/worksheets/_rels/sheet%u.xml.rels", i + 1);
            write_sheet_rels_buf(&b, wb->sheets[i]);
            ADD(rels_entry);
        }
    }

    /* 8. xl/workbook.xml */
    oxl_xmlbuf_cstr(&b,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    if (wb->date1904)
        oxl_xmlbuf_cstr(&b, "<workbookPr date1904=\"1\"/>");
    oxl_xmlbuf_cstr(&b, "<sheets>");
    for (uint32_t i = 0; i < wb->sheet_count; i++) {
        oxl_xmlbuf_cstr(&b, "<sheet name=\"");
        oxl_xmlbuf_attr_val(&b, wb->sheets[i]->name ? wb->sheets[i]->name : "Sheet");
        oxl_xmlbuf_cstr(&b, "\" sheetId=\"");
        oxl_xmlbuf_uint(&b, i + 1);
        oxl_xmlbuf_cstr(&b, "\" r:id=\"rId");
        oxl_xmlbuf_uint(&b, i + 2);
        oxl_xmlbuf_cstr(&b, "\"/>");
    }
    oxl_xmlbuf_cstr(&b, "</sheets>");
    if (wb->defined_name_count > 0) {
        oxl_xmlbuf_cstr(&b, "<definedNames>");
        for (uint32_t i = 0; i < wb->defined_name_count; i++) {
            const OxlDefinedName *dn = &wb->defined_names[i];
            oxl_xmlbuf_cstr(&b, "<definedName name=\"");
            oxl_xmlbuf_attr_val(&b, dn->name ? dn->name : "");
            if (dn->local_sheet_id >= 0) {
                oxl_xmlbuf_cstr(&b, "\" localSheetId=\"");
                oxl_xmlbuf_uint(&b, (uint32_t)dn->local_sheet_id);
            }
            if (dn->hidden) {
                oxl_xmlbuf_cstr(&b, "\" hidden=\"1");
            }
            oxl_xmlbuf_cstr(&b, "\">");
            oxl_xmlbuf_text(&b, dn->value ? dn->value : "");
            oxl_xmlbuf_cstr(&b, "</definedName>");
        }
        oxl_xmlbuf_cstr(&b, "</definedNames>");
    }
    oxl_xmlbuf_cstr(&b, "</workbook>");
    ADD("xl/workbook.xml");

done:
    oxl_xmlbuf_free(&b);
    oxl_zipw_close(zw);
    return rc;
}
