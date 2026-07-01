#pragma once
#include "cell.h"
#include "styles.h"
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

/* Phase 15: Sheet Protection */

typedef struct {
    char    *password_hash;
    char    *algorithm_name;
    char    *hash_value;
    char    *salt_value;
    uint32_t spin_count;
    uint8_t  sheet;
    uint8_t  objects;
    uint8_t  scenarios;
    uint8_t  format_cells;
    uint8_t  format_columns;
    uint8_t  format_rows;
    uint8_t  insert_columns;
    uint8_t  insert_rows;
    uint8_t  insert_hyperlinks;
    uint8_t  delete_columns;
    uint8_t  delete_rows;
    uint8_t  select_locked;
    uint8_t  sort;
    uint8_t  auto_filter;
    uint8_t  pivot_tables;
    uint8_t  select_unlocked;
    uint8_t  has_protection;
} OxlSheetProtection;

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

/* Phase 16: Conditional Formatting */

typedef struct {
    char    *type;    /* "min","max","num","percent","percentile","formula" */
    char    *val;
    uint32_t rgb;
    uint8_t  has_rgb;
} OxlCfvo;

typedef struct {
    char    *type;       /* "cellIs","expression","colorScale","dataBar","top10",
                            "aboveAverage","containsText","notContainsText",
                            "beginsWith","endsWith","duplicateValues","uniqueValues" */
    char    *operator_;
    char    *formula;
    char    *formula2;
    char    *text;
    int32_t  priority;
    uint8_t  stop_if_true;
    OxlFontDef  *font;   /* NULL = no font override */
    OxlFillDef  *fill;   /* NULL = no fill override */
    OxlBorderDef *border; /* NULL = no border override */
    int32_t  dxf_id;     /* -1 = unset */
    uint8_t  top10_top;
    uint8_t  top10_percent;
    uint32_t top10_rank;
    uint8_t  above_avg;
    uint8_t  equal_avg;
    OxlCfvo  cfvos[3];
    uint32_t cfvo_count;
    uint32_t colors[3];  /* ARGB */
    uint32_t color_count;
    uint8_t  data_bar_show_value;
} OxlCfRule;

typedef struct {
    char      *sqref;
    OxlCfRule *rules;
    uint32_t   rule_count;
    uint32_t   rule_cap;
} OxlCf;

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

    /* Phase 15: Sheet Protection */
    OxlSheetProtection protection;

    /* Phase 16: Conditional Formatting */
    OxlCf    *cond_fmts;
    uint32_t  cf_count;
    uint32_t  cf_cap;
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

void oxl_cfvo_free_fields(OxlCfvo *v);
void oxl_cf_rule_free_fields(OxlCfRule *rule);
int  oxl_worksheet_add_cf_rule(OxlWorksheet *ws, const char *sqref, const OxlCfRule *rule);
