#include "xml_sheet_rels.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>

void oxl_hyperlink_rels_init(OxlHyperlinkRels *h) {
    h->entries = NULL; h->count = 0; h->cap = 0;
}

void oxl_hyperlink_rels_free(OxlHyperlinkRels *h) {
    for (uint32_t i = 0; i < h->count; i++) {
        free(h->entries[i].rid);
        free(h->entries[i].url);
    }
    free(h->entries);
    oxl_hyperlink_rels_init(h);
}

const char *oxl_hyperlink_rels_lookup(const OxlHyperlinkRels *h, const char *rid) {
    if (!rid) return NULL;
    for (uint32_t i = 0; i < h->count; i++)
        if (h->entries[i].rid && strcmp(h->entries[i].rid, rid) == 0)
            return h->entries[i].url;
    return NULL;
}

static const char *attr(const char **attrs, const char *key) {
    for (; *attrs; attrs += 2) if (strcmp(attrs[0], key) == 0) return attrs[1];
    return NULL;
}

static void XMLCALL rels_start(void *ud, const char *name, const char **attrs) {
    OxlHyperlinkRels *h = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;
    if (strcmp(name, "Relationship") != 0) return;

    const char *type   = attr(attrs, "Type");
    const char *id     = attr(attrs, "Id");
    const char *target = attr(attrs, "Target");
    if (!type || !id || !target) return;

    /* Only capture hyperlink relationships */
    size_t tlen = strlen(type);
    const char *suffix = "/hyperlink";
    size_t slen = strlen(suffix);
    if (tlen < slen || strcmp(type + tlen - slen, suffix) != 0) return;

    /* Grow array */
    if (h->count >= h->cap) {
        uint32_t newcap = h->cap ? h->cap * 2 : 4;
        OxlHyperlinkEntry *p = realloc(h->entries, newcap * sizeof(*h->entries));
        if (!p) return;
        h->entries = p;
        h->cap = newcap;
    }
    h->entries[h->count].rid = strdup(id);
    h->entries[h->count].url = strdup(target);
    h->count++;
}

int oxl_parse_sheet_rels(const char *buf, size_t len, OxlHyperlinkRels *h) {
    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, h);
    XML_SetStartElementHandler(p, rels_start);
    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);
    return ok ? 0 : -1;
}
