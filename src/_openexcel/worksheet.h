#pragma once
#include "cell.h"
#include <stdint.h>

/* ── Feature A: Column & Row Dimensions ─────────────────────────────────── */

typedef struct {
    uint16_t col_min;      /* 1-based, inclusive */
    uint16_t col_max;      /* 1-based, inclusive */
    double   width;        /* character units; 0.0 = not set */
    int      hidden;
    int      best_fit;
    int      custom_width;
} OxlColDim;

typedef struct {
    uint32_t row_idx;      /* 1-based */
    double   height;       /* points; 0.0 = not set */
    int      hidden;
    int      custom_height;
} OxlRowDim;

/* ── Feature B: Merged Cells ─────────────────────────────────────────────── */

typedef struct {
    uint32_t min_row;  /* 1-based */
    uint16_t min_col;  /* 1-based */
    uint32_t max_row;  /* 1-based */
    uint16_t max_col;  /* 1-based */
} OxlMergedCell;

/* ── Worksheet ───────────────────────────────────────────────────────────── */

typedef struct {
    char    *name;              /* sheet display name (UTF-8, owned) */
    char    *rel_path;          /* ZIP entry path, e.g. "xl/worksheets/sheet1.xml" */
    OxlCell *cells;             /* flat array, document/row-major order */
    uint32_t cell_count;
    uint32_t cell_capacity;
    uint32_t row_count;         /* max row index seen + 1 */
    uint32_t col_count;         /* max col index seen + 1 */

    /* Feature A */
    OxlColDim  *col_dims;
    uint32_t    col_dim_count;
    uint32_t    col_dim_capacity;
    OxlRowDim  *row_dims;
    uint32_t    row_dim_count;
    uint32_t    row_dim_capacity;

    /* Feature B */
    OxlMergedCell *merged_cells;
    uint32_t       merged_cell_count;
    uint32_t       merged_cell_capacity;

    /* Feature C: Sheet Views / Freeze Panes */
    char    *freeze_panes;    /* e.g. "B2"; NULL if none (heap-allocated) */
    int      zoom_scale;      /* 0 = default (100%) */
    int      show_gridlines;  /* 1 = show (default), 0 = hide */

    /* Feature D: Tab Color */
    char tab_color[9];        /* AARRGGBB or RRGGBB hex, empty string if none */

    /* Feature E: Auto-Filter */
    char *auto_filter_ref;    /* e.g. "A1:D1"; NULL if none (heap-allocated) */
} OxlWorksheet;

OxlWorksheet *oxl_worksheet_new(const char *name, const char *rel_path);
void          oxl_worksheet_free(OxlWorksheet *ws);

/* Append a cell (copied by value). Caller must have already set all fields. */
int           oxl_worksheet_append_cell(OxlWorksheet *ws, OxlCell *c);

/* Feature A helpers */
int oxl_worksheet_set_col_dim(OxlWorksheet *ws, uint16_t col_min, uint16_t col_max,
                               double width, int hidden, int best_fit, int custom_width);
int oxl_worksheet_set_row_dim(OxlWorksheet *ws, uint32_t row_idx, double height,
                               int hidden, int custom_height);

/* Feature B helper */
int oxl_worksheet_add_merge(OxlWorksheet *ws, uint32_t min_row, uint16_t min_col,
                             uint32_t max_row, uint16_t max_col);
