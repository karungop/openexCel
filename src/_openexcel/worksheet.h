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

/* Phase 13: Data Validation */

typedef struct {
    char    *type;           /* "list", "whole", "decimal", "date", "time", "textLength", "custom" */
    char    *dv_operator;    /* "between", "notBetween", "equal", "notEqual", "greaterThan", etc. */
    char    *formula1;       /* first formula/value */
    char    *formula2;       /* second formula/value (for range checks) */
    char    *sqref;          /* cell range, e.g. "A1:A10" */
    char    *error_message;
    char    *error_title;
    char    *error_style;    /* "stop", "warning", "information" */
    char    *prompt_message;
    char    *prompt_title;
    uint8_t  allow_blank;
    uint8_t  show_drop_down;       /* 1 = hide dropdown (XML attr "showDropDown"="1" means hidden) */
    uint8_t  show_error_message;
    uint8_t  show_input_message;
} OxlDataValidation;

/* Phase 14: Page Setup & Print Options */

typedef struct {
    char    *orientation;   /* "portrait", "landscape"; NULL = not set */
    uint32_t paper_size;    /* 1=Letter, 9=A4, etc.; 0 = not set */
    uint32_t scale;         /* 10–400; 0 = not set */
    uint32_t fit_to_width;  /* pages wide (fit to page mode) */
    uint32_t fit_to_height; /* pages tall (fit to page mode) */
    uint8_t  fit_to_page;   /* 0 = scale mode, 1 = fit-to-page mode */
    uint8_t  has_setup;     /* 1 = at least one field was set */
} OxlPageSetup;

typedef struct {
    double   left;
    double   right;
    double   top;
    double   bottom;
    double   header;
    double   footer;
    uint8_t  has_margins;  /* 1 = explicit margins set */
} OxlPageMargins;

typedef struct {
    uint8_t  grid_lines;           /* print grid lines */
    uint8_t  headings;             /* print row/col headings */
    uint8_t  horizontal_centered;  /* center on page horizontally */
    uint8_t  vertical_centered;    /* center on page vertically */
    uint8_t  has_options;          /* 1 = at least one flag set */
} OxlPrintOptions;

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

    /* Phase 13: Data Validations */
    OxlDataValidation *data_validations;
    uint32_t           dv_count;
    uint32_t           dv_cap;

    /* Phase 14: Page Setup & Print Options */
    OxlPageSetup    page_setup;
    OxlPageMargins  page_margins;
    OxlPrintOptions print_options;
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

/* Phase 13: Data Validation helpers */
void oxl_data_validation_free_fields(OxlDataValidation *dv);
int  oxl_worksheet_add_data_validation(OxlWorksheet *ws, const OxlDataValidation *dv);

/* Phase 14: no extra helper functions needed — fields are embedded structs */
