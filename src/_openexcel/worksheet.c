#include "worksheet.h"
#include <stdlib.h>
#include <string.h>

void oxl_data_validation_free_fields(OxlDataValidation *dv) {
    free(dv->type);
    free(dv->dv_operator);
    free(dv->formula1);
    free(dv->formula2);
    free(dv->sqref);
    free(dv->error_message);
    free(dv->error_title);
    free(dv->error_style);
    free(dv->prompt_message);
    free(dv->prompt_title);
}

OxlWorksheet *oxl_worksheet_new(const char *name, const char *rel_path) {
    OxlWorksheet *ws = calloc(1, sizeof(OxlWorksheet));
    if (!ws) return NULL;
    ws->name = name ? strdup(name) : NULL;
    ws->rel_path = rel_path ? strdup(rel_path) : NULL;
    ws->show_gridlines = 1;  /* Feature C: default show gridlines */
    /* Phase 14: default page margins (Excel defaults) */
    ws->page_margins.left   = 0.7;
    ws->page_margins.right  = 0.7;
    ws->page_margins.top    = 0.75;
    ws->page_margins.bottom = 0.75;
    ws->page_margins.header = 0.3;
    ws->page_margins.footer = 0.3;
    /* has_margins stays 0 — only set to 1 when user explicitly sets margins */
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
    /* Feature A */
    free(ws->col_dims);
    free(ws->row_dims);
    /* Feature B */
    free(ws->merged_cells);
    /* Feature C */
    free(ws->freeze_panes);
    /* Feature E */
    free(ws->auto_filter_ref);
    /* Phase 13: data validations */
    for (uint32_t i = 0; i < ws->dv_count; i++)
        oxl_data_validation_free_fields(&ws->data_validations[i]);
    free(ws->data_validations);
    /* Phase 14: page setup orientation is heap-allocated */
    free(ws->page_setup.orientation);
    /* Phase 15: sheet protection string fields */
    free(ws->protection.password_hash);
    free(ws->protection.algorithm_name);
    free(ws->protection.hash_value);
    free(ws->protection.salt_value);
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

/* ── Feature A: Column Dimensions ────────────────────────────────────────── */

int oxl_worksheet_set_col_dim(OxlWorksheet *ws, uint16_t col_min, uint16_t col_max,
                               double width, int hidden, int best_fit, int custom_width) {
    /* Update existing entry if found */
    for (uint32_t i = 0; i < ws->col_dim_count; i++) {
        if (ws->col_dims[i].col_min == col_min) {
            ws->col_dims[i].col_max      = col_max;
            ws->col_dims[i].width        = width;
            ws->col_dims[i].hidden       = hidden;
            ws->col_dims[i].best_fit     = best_fit;
            ws->col_dims[i].custom_width = custom_width;
            return 0;
        }
    }
    /* Append new entry */
    if (ws->col_dim_count == ws->col_dim_capacity) {
        uint32_t cap = ws->col_dim_capacity ? ws->col_dim_capacity * 2 : 16;
        OxlColDim *p = realloc(ws->col_dims, cap * sizeof(OxlColDim));
        if (!p) return -1;
        ws->col_dims = p;
        ws->col_dim_capacity = cap;
    }
    OxlColDim *d = &ws->col_dims[ws->col_dim_count++];
    d->col_min      = col_min;
    d->col_max      = col_max;
    d->width        = width;
    d->hidden       = hidden;
    d->best_fit     = best_fit;
    d->custom_width = custom_width;
    return 0;
}

/* ── Feature A: Row Dimensions ───────────────────────────────────────────── */

int oxl_worksheet_set_row_dim(OxlWorksheet *ws, uint32_t row_idx, double height,
                               int hidden, int custom_height) {
    /* Update existing entry if found */
    for (uint32_t i = 0; i < ws->row_dim_count; i++) {
        if (ws->row_dims[i].row_idx == row_idx) {
            ws->row_dims[i].height        = height;
            ws->row_dims[i].hidden        = hidden;
            ws->row_dims[i].custom_height = custom_height;
            return 0;
        }
    }
    /* Append new entry */
    if (ws->row_dim_count == ws->row_dim_capacity) {
        uint32_t cap = ws->row_dim_capacity ? ws->row_dim_capacity * 2 : 16;
        OxlRowDim *p = realloc(ws->row_dims, cap * sizeof(OxlRowDim));
        if (!p) return -1;
        ws->row_dims = p;
        ws->row_dim_capacity = cap;
    }
    OxlRowDim *d = &ws->row_dims[ws->row_dim_count++];
    d->row_idx      = row_idx;
    d->height       = height;
    d->hidden       = hidden;
    d->custom_height = custom_height;
    return 0;
}

/* ── Feature B: Merged Cells ─────────────────────────────────────────────── */

int oxl_worksheet_add_merge(OxlWorksheet *ws, uint32_t min_row, uint16_t min_col,
                             uint32_t max_row, uint16_t max_col) {
    if (ws->merged_cell_count == ws->merged_cell_capacity) {
        uint32_t cap = ws->merged_cell_capacity ? ws->merged_cell_capacity * 2 : 8;
        OxlMergedCell *p = realloc(ws->merged_cells, cap * sizeof(OxlMergedCell));
        if (!p) return -1;
        ws->merged_cells = p;
        ws->merged_cell_capacity = cap;
    }
    OxlMergedCell *m = &ws->merged_cells[ws->merged_cell_count++];
    m->min_row = min_row;
    m->min_col = min_col;
    m->max_row = max_row;
    m->max_col = max_col;
    return 0;
}

/* ── Phase 13: Data Validations ──────────────────────────────────────────── */

int oxl_worksheet_add_data_validation(OxlWorksheet *ws, const OxlDataValidation *dv) {
    if (ws->dv_count == ws->dv_cap) {
        uint32_t cap = ws->dv_cap ? ws->dv_cap * 2 : 8;
        OxlDataValidation *p = realloc(ws->data_validations, cap * sizeof(OxlDataValidation));
        if (!p) return -1;
        ws->data_validations = p;
        ws->dv_cap = cap;
    }
    OxlDataValidation *dst = &ws->data_validations[ws->dv_count++];
    memset(dst, 0, sizeof(*dst));
    dst->type               = dv->type           ? strdup(dv->type)           : NULL;
    dst->dv_operator        = dv->dv_operator     ? strdup(dv->dv_operator)     : NULL;
    dst->formula1           = dv->formula1        ? strdup(dv->formula1)        : NULL;
    dst->formula2           = dv->formula2        ? strdup(dv->formula2)        : NULL;
    dst->sqref              = dv->sqref           ? strdup(dv->sqref)           : NULL;
    dst->error_message      = dv->error_message   ? strdup(dv->error_message)   : NULL;
    dst->error_title        = dv->error_title     ? strdup(dv->error_title)     : NULL;
    dst->error_style        = dv->error_style     ? strdup(dv->error_style)     : NULL;
    dst->prompt_message     = dv->prompt_message  ? strdup(dv->prompt_message)  : NULL;
    dst->prompt_title       = dv->prompt_title    ? strdup(dv->prompt_title)    : NULL;
    dst->allow_blank        = dv->allow_blank;
    dst->show_drop_down     = dv->show_drop_down;
    dst->show_error_message = dv->show_error_message;
    dst->show_input_message = dv->show_input_message;
    return 0;
}
