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
    /* Phase 16: conditional formatting */
    if (ws->cond_fmts) {
        for (uint32_t i = 0; i < ws->cf_count; i++) {
            OxlCf *cf = &ws->cond_fmts[i];
            free(cf->sqref);
            for (uint32_t j = 0; j < cf->rule_count; j++)
                oxl_cf_rule_free_fields(&cf->rules[j]);
            free(cf->rules);
        }
        free(ws->cond_fmts);
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

/* ── Phase 16: Conditional Formatting ───────────────────────────────────── */

void oxl_cfvo_free_fields(OxlCfvo *v) {
    free(v->type);
    free(v->val);
}

void oxl_cf_rule_free_fields(OxlCfRule *rule) {
    free(rule->type);
    free(rule->operator_);
    free(rule->formula);
    free(rule->formula2);
    free(rule->text);
    /* Note: font/fill/border are NOT owned by OxlCfRule (they point to rule_obj Python data).
       After dxf_id is assigned, the pointers are only used during the pre-pass in writer.c.
       For heap-allocated copies (from Python add_conditional_formatting), we need to free them.
       The convention: if dxf_id >= 0, the font/fill/border were used for DXF creation.
       We free them here since oxl_worksheet_add_cf_rule makes copies. */
    free(rule->font);
    free(rule->fill);
    free(rule->border);
    for (uint32_t i = 0; i < rule->cfvo_count; i++)
        oxl_cfvo_free_fields(&rule->cfvos[i]);
}

int oxl_worksheet_add_cf_rule(OxlWorksheet *ws, const char *sqref, const OxlCfRule *rule) {
    /* Find or create an OxlCf block for this sqref */
    OxlCf *cf = NULL;
    for (uint32_t i = 0; i < ws->cf_count; i++) {
        if (ws->cond_fmts[i].sqref && sqref &&
            strcmp(ws->cond_fmts[i].sqref, sqref) == 0) {
            cf = &ws->cond_fmts[i];
            break;
        }
    }
    if (!cf) {
        /* Create new OxlCf block */
        if (ws->cf_count >= ws->cf_cap) {
            uint32_t cap = ws->cf_cap ? ws->cf_cap * 2 : 8;
            OxlCf *p = realloc(ws->cond_fmts, cap * sizeof(OxlCf));
            if (!p) return -1;
            ws->cond_fmts = p;
            ws->cf_cap = cap;
        }
        cf = &ws->cond_fmts[ws->cf_count++];
        memset(cf, 0, sizeof(*cf));
        cf->sqref = sqref ? strdup(sqref) : NULL;
    }
    /* Append rule to this OxlCf block */
    if (cf->rule_count >= cf->rule_cap) {
        uint32_t cap = cf->rule_cap ? cf->rule_cap * 2 : 4;
        OxlCfRule *p = realloc(cf->rules, cap * sizeof(OxlCfRule));
        if (!p) return -1;
        cf->rules = p;
        cf->rule_cap = cap;
    }
    OxlCfRule *dst = &cf->rules[cf->rule_count++];
    memset(dst, 0, sizeof(*dst));
    dst->type      = rule->type      ? strdup(rule->type)      : NULL;
    dst->operator_ = rule->operator_ ? strdup(rule->operator_) : NULL;
    dst->formula   = rule->formula   ? strdup(rule->formula)   : NULL;
    dst->formula2  = rule->formula2  ? strdup(rule->formula2)  : NULL;
    dst->text      = rule->text      ? strdup(rule->text)      : NULL;
    dst->priority  = rule->priority;
    dst->stop_if_true = rule->stop_if_true;
    dst->dxf_id    = rule->dxf_id;
    dst->top10_top = rule->top10_top;
    dst->top10_percent = rule->top10_percent;
    dst->top10_rank = rule->top10_rank;
    dst->above_avg  = rule->above_avg;
    dst->equal_avg  = rule->equal_avg;
    dst->cfvo_count = rule->cfvo_count;
    dst->color_count = rule->color_count;
    dst->data_bar_show_value = rule->data_bar_show_value;
    /* Copy cfvos */
    for (uint32_t i = 0; i < rule->cfvo_count && i < 3; i++) {
        dst->cfvos[i].type = rule->cfvos[i].type ? strdup(rule->cfvos[i].type) : NULL;
        dst->cfvos[i].val  = rule->cfvos[i].val  ? strdup(rule->cfvos[i].val)  : NULL;
        dst->cfvos[i].rgb  = rule->cfvos[i].rgb;
        dst->cfvos[i].has_rgb = rule->cfvos[i].has_rgb;
    }
    /* Copy colors */
    for (uint32_t i = 0; i < rule->color_count && i < 3; i++)
        dst->colors[i] = rule->colors[i];
    /* Deep-copy font/fill/border if present */
    dst->font = NULL;
    dst->fill = NULL;
    dst->border = NULL;
    if (rule->font) {
        dst->font = calloc(1, sizeof(OxlFontDef));
        if (dst->font) {
            *dst->font = *rule->font;
            dst->font->name = rule->font->name ? strdup(rule->font->name) : NULL;
        }
    }
    if (rule->fill) {
        dst->fill = calloc(1, sizeof(OxlFillDef));
        if (dst->fill) {
            *dst->fill = *rule->fill;
            dst->fill->pattern_type = rule->fill->pattern_type ? strdup(rule->fill->pattern_type) : NULL;
        }
    }
    if (rule->border) {
        dst->border = calloc(1, sizeof(OxlBorderDef));
        if (dst->border) {
            *dst->border = *rule->border;
            dst->border->left.style     = rule->border->left.style     ? strdup(rule->border->left.style)     : NULL;
            dst->border->right.style    = rule->border->right.style    ? strdup(rule->border->right.style)    : NULL;
            dst->border->top.style      = rule->border->top.style      ? strdup(rule->border->top.style)      : NULL;
            dst->border->bottom.style   = rule->border->bottom.style   ? strdup(rule->border->bottom.style)   : NULL;
            dst->border->diagonal.style = rule->border->diagonal.style ? strdup(rule->border->diagonal.style) : NULL;
        }
    }
    return 0;
}
