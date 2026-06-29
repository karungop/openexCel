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
    if (wb->defined_names) {
        for (uint32_t i = 0; i < wb->defined_name_count; i++) {
            free(wb->defined_names[i].name);
            free(wb->defined_names[i].value);
        }
        free(wb->defined_names);
    }
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

int oxl_workbook_add_defined_name(OxlWorkbook *wb, const char *name,
                                   const char *value, int local_sheet_id, int hidden) {
    /* Check for existing entry with same name and local_sheet_id → update */
    for (uint32_t i = 0; i < wb->defined_name_count; i++) {
        OxlDefinedName *dn = &wb->defined_names[i];
        if (dn->local_sheet_id == local_sheet_id &&
                dn->name && strcmp(dn->name, name) == 0) {
            char *new_val = value ? strdup(value) : NULL;
            if (value && !new_val) return -1;
            free(dn->value);
            dn->value  = new_val;
            dn->hidden = hidden;
            return 0;
        }
    }
    /* Grow array if needed */
    if (wb->defined_name_count == wb->defined_name_capacity) {
        uint32_t cap = wb->defined_name_capacity ? wb->defined_name_capacity * 2 : 8;
        OxlDefinedName *p = realloc(wb->defined_names, cap * sizeof(OxlDefinedName));
        if (!p) return -1;
        wb->defined_names = p;
        wb->defined_name_capacity = cap;
    }
    OxlDefinedName *dn = &wb->defined_names[wb->defined_name_count];
    dn->name          = name  ? strdup(name)  : NULL;
    dn->value         = value ? strdup(value) : NULL;
    dn->local_sheet_id = local_sheet_id;
    dn->hidden        = hidden;
    if ((name && !dn->name) || (value && !dn->value)) {
        free(dn->name);
        free(dn->value);
        dn->name = dn->value = NULL;
        return -1;
    }
    wb->defined_name_count++;
    return 0;
}
