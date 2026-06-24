#include "workbook.h"
#include <stdlib.h>
#include <string.h>

OxlWorkbook *oxl_workbook_new(void) {
    OxlWorkbook *wb = calloc(1, sizeof(OxlWorkbook));
    if (!wb) return NULL;
    oxl_sst_init(&wb->sst);
    oxl_styles_init(&wb->styles);
    return wb;
}

void oxl_workbook_free(OxlWorkbook *wb) {
    if (!wb) return;
    for (uint32_t i = 0; i < wb->sheet_count; i++) oxl_worksheet_free(wb->sheets[i]);
    free(wb->sheets);
    oxl_sst_free(&wb->sst);
    oxl_styles_free(&wb->styles);
    free(wb);
}

int oxl_workbook_add_sheet(OxlWorkbook *wb, OxlWorksheet *ws) {
    if (wb->sheet_count == wb->sheet_capacity) {
        uint32_t cap = wb->sheet_capacity ? wb->sheet_capacity * 2 : 8;
        OxlWorksheet **p = realloc(wb->sheets, cap * sizeof(OxlWorksheet *));
        if (!p) return -1;
        wb->sheets = p;
        wb->sheet_capacity = cap;
    }
    int idx = (int)wb->sheet_count;
    wb->sheets[wb->sheet_count++] = ws;
    return idx;
}

OxlWorksheet *oxl_workbook_get_sheet_by_name(OxlWorkbook *wb, const char *name) {
    for (uint32_t i = 0; i < wb->sheet_count; i++) {
        if (wb->sheets[i]->name && strcmp(wb->sheets[i]->name, name) == 0)
            return wb->sheets[i];
    }
    return NULL;
}

OxlWorksheet *oxl_workbook_get_sheet_by_index(OxlWorkbook *wb, uint32_t idx) {
    if (idx >= wb->sheet_count) return NULL;
    return wb->sheets[idx];
}
