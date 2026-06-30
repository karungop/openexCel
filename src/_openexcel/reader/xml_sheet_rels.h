#pragma once
#include <stddef.h>
#include <stdint.h>

/* One rId → URL mapping entry */
typedef struct {
    char *rid;   /* e.g. "rId1" */
    char *url;   /* target URL or anchor */
} OxlHyperlinkEntry;

/* Collection of rId→URL mappings parsed from a sheet rels file */
typedef struct {
    OxlHyperlinkEntry *entries;
    uint32_t           count;
    uint32_t           cap;
} OxlHyperlinkRels;

void        oxl_hyperlink_rels_init(OxlHyperlinkRels *h);
void        oxl_hyperlink_rels_free(OxlHyperlinkRels *h);
const char *oxl_hyperlink_rels_lookup(const OxlHyperlinkRels *h, const char *rid);

/* Parse xl/worksheets/_rels/sheetN.xml.rels into h.
   Returns 0 on success, -1 on error. h must already be init'd. */
int oxl_parse_sheet_rels(const char *buf, size_t len, OxlHyperlinkRels *h);
