#include "worksheet.h"
#include <stdlib.h>
#include <string.h>

OxlWorksheet *oxl_worksheet_new(const char *name, const char *rel_path) {
    OxlWorksheet *ws = calloc(1, sizeof(OxlWorksheet));
    if (!ws) return NULL;
    ws->name = name ? strdup(name) : NULL;
    ws->rel_path = rel_path ? strdup(rel_path) : NULL;
    ws->show_gridlines = 1;  /* Feature C: default show gridlines */
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
