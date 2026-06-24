#include "worksheet.h"
#include <stdlib.h>
#include <string.h>

OxlWorksheet *oxl_worksheet_new(const char *name, const char *rel_path) {
    OxlWorksheet *ws = calloc(1, sizeof(OxlWorksheet));
    if (!ws) return NULL;
    ws->name = name ? strdup(name) : NULL;
    ws->rel_path = rel_path ? strdup(rel_path) : NULL;
    return ws;
}

void oxl_worksheet_free(OxlWorksheet *ws) {
    if (!ws) return;
    free(ws->name);
    free(ws->rel_path);
    if (ws->cells) {
        for (uint32_t i = 0; i < ws->cell_count; i++) oxl_cell_free(&ws->cells[i]);
        free(ws->cells);
    }
    free(ws);
}

int oxl_worksheet_append_cell(OxlWorksheet *ws, OxlCell *c) {
    if (ws->cell_count == ws->cell_capacity) {
        uint32_t cap = ws->cell_capacity ? ws->cell_capacity * 2 : 1024;
        OxlCell *p = realloc(ws->cells, cap * sizeof(OxlCell));
        if (!p) return -1;
        ws->cells = p;
        ws->cell_capacity = cap;
    }
    ws->cells[ws->cell_count++] = *c;
    if (c->row + 1 > ws->row_count) ws->row_count = c->row + 1;
    if (c->col + 1 > ws->col_count) ws->col_count = c->col + 1;
    return 0;
}
