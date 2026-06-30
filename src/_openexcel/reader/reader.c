#include "reader.h"
#include "zip_reader.h"
#include "xml_workbook.h"
#include "xml_sst.h"
#include "xml_styles.h"
#include "xml_sheet.h"
#include "xml_sheet_rels.h"
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

#define ERR(fmt, ...) \
    do { if (err_buf) snprintf(err_buf, err_cap, fmt, ##__VA_ARGS__); } while(0)

OxlWorkbook *oxl_read_workbook(const char *path, char *err_buf, size_t err_cap) {
    OxlZipReader *zr = oxl_zip_open(path);
    if (!zr) { ERR("Cannot open file: %s", path); return NULL; }

    OxlWorkbook *wb = oxl_workbook_new();
    if (!wb) { oxl_zip_close(zr); ERR("Out of memory"); return NULL; }

    char *buf = NULL;
    size_t sz = 0;

#define EXTRACT(entry) \
    do { free(buf); buf = NULL; \
         if (oxl_zip_extract(zr, entry, &buf, &sz) != 0) { \
             ERR("Missing required entry: %s", entry); goto fail; } } while(0)

#define EXTRACT_OPT(entry) \
    do { free(buf); buf = NULL; \
         oxl_zip_extract(zr, entry, &buf, &sz); } while(0)

    /* 1. Workbook relationships: rId → target path */
    EXTRACT("xl/_rels/workbook.xml.rels");
    if (oxl_parse_workbook_rels(buf, sz, wb) != 0) {
        ERR("Failed to parse workbook.xml.rels"); goto fail;
    }

    /* 2. Workbook: sheet names, date1904 — adds sheets with rId in rel_path temporarily */
    EXTRACT("xl/workbook.xml");
    if (oxl_parse_workbook_xml(buf, sz, wb) != 0) {
        ERR("Failed to parse workbook.xml"); goto fail;
    }

    /* After parsing workbook.xml, re-run rels to resolve rId → actual path.
       workbook.xml parser stores rIds in rel_path; rels parser replaces them with paths. */
    EXTRACT("xl/_rels/workbook.xml.rels");
    if (oxl_parse_workbook_rels(buf, sz, wb) != 0) {
        ERR("Failed to parse workbook rels"); goto fail;
    }

    /* 3. Styles (optional — needed for date detection) */
    EXTRACT_OPT("xl/styles.xml");
    if (buf && oxl_parse_styles(buf, sz, &wb->styles) != 0) {
        /* Non-fatal: just means no date detection */
    }

    /* 4. Shared strings (optional) */
    EXTRACT_OPT("xl/sharedStrings.xml");
    if (buf && oxl_parse_sst(buf, sz, &wb->sst) != 0) {
        ERR("Failed to parse sharedStrings.xml"); goto fail;
    }

    /* 5. Parse each worksheet */
    for (uint32_t i = 0; i < wb->sheet_count; i++) {
        OxlWorksheet *ws = wb->sheets[i];
        if (!ws->rel_path) { ERR("Sheet %u has no path", i); goto fail; }

        free(buf); buf = NULL;
        if (oxl_zip_extract(zr, ws->rel_path, &buf, &sz) != 0) {
            ERR("Cannot extract sheet: %s", ws->rel_path); goto fail;
        }

        /* Try to parse sheet rels file for hyperlinks.
           Derive rels path: "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels" */
        OxlHyperlinkRels sheet_rels;
        oxl_hyperlink_rels_init(&sheet_rels);
        {
            const char *rel_path = ws->rel_path;
            const char *last_slash = strrchr(rel_path, '/');
            const char *fname = last_slash ? last_slash + 1 : rel_path;
            char rels_path[256];
            if (last_slash) {
                size_t dir_len = (size_t)(last_slash - rel_path);
                snprintf(rels_path, sizeof(rels_path), "%.*s/_rels/%s.rels",
                         (int)dir_len, rel_path, fname);
            } else {
                snprintf(rels_path, sizeof(rels_path), "_rels/%s.rels", fname);
            }
            char *rels_buf = NULL; size_t rels_sz = 0;
            oxl_zip_extract(zr, rels_path, &rels_buf, &rels_sz);
            if (rels_buf) {
                oxl_parse_sheet_rels(rels_buf, rels_sz, &sheet_rels);
                free(rels_buf);
            }
        }

        if (oxl_parse_sheet(buf, sz, ws, wb, &sheet_rels) != 0) {
            oxl_hyperlink_rels_free(&sheet_rels);
            ERR("Failed to parse sheet: %s", ws->rel_path); goto fail;
        }
        oxl_hyperlink_rels_free(&sheet_rels);
    }

    free(buf);
    oxl_zip_close(zr);
    return wb;

fail:
    free(buf);
    oxl_zip_close(zr);
    oxl_workbook_free(wb);
    return NULL;
}
