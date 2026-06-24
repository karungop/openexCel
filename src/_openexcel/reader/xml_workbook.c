#include "xml_workbook.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

/* ---- workbook.xml parser ---- */

typedef struct {
    OxlWorkbook *wb;
    int          in_sheets;
    /* rId → sheet order: we store rId strings alongside sheet objects temporarily */
    char        **rids;     /* rids[i] = strdup of rId for sheet i */
    uint32_t      rid_count;
    uint32_t      rid_cap;
} WbCtx;

static const char *attr(const char **attrs, const char *key) {
    for (; *attrs; attrs += 2) if (strcmp(attrs[0], key) == 0) return attrs[1];
    return NULL;
}

static void XMLCALL wb_start(void *ud, const char *name, const char **attrs) {
    WbCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "workbookPr") == 0) {
        const char *d = attr(attrs, "date1904");
        if (d && (strcmp(d,"1") == 0 || strcmp(d,"true") == 0)) c->wb->date1904 = 1;
    } else if (strcmp(name, "sheets") == 0) {
        c->in_sheets = 1;
    } else if (strcmp(name, "sheet") == 0 && c->in_sheets) {
        const char *sheet_name = attr(attrs, "name");
        const char *rid        = attr(attrs, "r:id");
        if (!rid) rid = attr(attrs, "id"); /* fallback */
        OxlWorksheet *ws = oxl_worksheet_new(sheet_name, NULL);
        if (!ws) return;
        int idx = oxl_workbook_add_sheet(c->wb, ws);
        if (idx < 0) { oxl_worksheet_free(ws); return; }
        /* remember rId for this sheet */
        if (c->rid_count == c->rid_cap) {
            uint32_t cap = c->rid_cap ? c->rid_cap * 2 : 8;
            char **p = realloc(c->rids, cap * sizeof(char *));
            if (!p) return;
            c->rids = p;
            c->rid_cap = cap;
        }
        c->rids[c->rid_count++] = rid ? strdup(rid) : NULL;
    }
}

static void XMLCALL wb_end(void *ud, const char *name) {
    WbCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;
    if (strcmp(name, "sheets") == 0) c->in_sheets = 0;
}

int oxl_parse_workbook_xml(const char *buf, size_t len, OxlWorkbook *wb) {
    WbCtx c = {0};
    c.wb = wb;

    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, &c);
    XML_SetElementHandler(p, wb_start, wb_end);

    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);

    /* Store rIds temporarily in rel_path (will be overwritten by rels parser) */
    for (uint32_t i = 0; i < c.rid_count && i < wb->sheet_count; i++) {
        free(wb->sheets[i]->rel_path);
        wb->sheets[i]->rel_path = c.rids[i]; /* ownership transferred */
    }
    free(c.rids);
    return ok ? 0 : -1;
}

/* ---- workbook.xml.rels parser ---- */

typedef struct {
    OxlWorkbook *wb;
} RelsCtx;

static void XMLCALL rels_start(void *ud, const char *name, const char **attrs) {
    RelsCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "Relationship") != 0) return;

    const char *id     = attr(attrs, "Id");
    const char *type   = attr(attrs, "Type");
    const char *target = attr(attrs, "Target");
    if (!id || !type || !target) return;

    /* We only care about worksheet relationships */
    if (!strstr(type, "worksheet")) return;

    /* Find the sheet whose rel_path currently holds this rId (from workbook.xml) */
    for (uint32_t i = 0; i < c->wb->sheet_count; i++) {
        OxlWorksheet *ws = c->wb->sheets[i];
        if (ws->rel_path && strcmp(ws->rel_path, id) == 0) {
            /* Build full path: if Target starts with '/', use as-is; else prepend "xl/" */
            char full[512];
            if (target[0] == '/') {
                snprintf(full, sizeof(full), "%s", target + 1);
            } else {
                snprintf(full, sizeof(full), "xl/%s", target);
            }
            free(ws->rel_path);
            ws->rel_path = strdup(full);
            break;
        }
    }
}

int oxl_parse_workbook_rels(const char *buf, size_t len, OxlWorkbook *wb) {
    RelsCtx c = {.wb = wb};
    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, &c);
    XML_SetElementHandler(p, rels_start, NULL);
    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);
    return ok ? 0 : -1;
}
