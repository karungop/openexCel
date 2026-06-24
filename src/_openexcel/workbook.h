#pragma once
#include "worksheet.h"
#include "string_table.h"
#include "styles.h"

typedef struct {
    OxlWorksheet   **sheets;
    uint32_t         sheet_count;
    uint32_t         sheet_capacity;
    OxlStringTable   sst;
    OxlStyles        styles;
    int              date1904;  /* 0 = 1900 epoch (default), 1 = 1904 epoch */
} OxlWorkbook;

OxlWorkbook *oxl_workbook_new(void);
void         oxl_workbook_free(OxlWorkbook *wb);

/* Add a sheet; workbook takes ownership. Returns sheet index or -1 on error. */
int          oxl_workbook_add_sheet(OxlWorkbook *wb, OxlWorksheet *ws);

OxlWorksheet *oxl_workbook_get_sheet_by_name(OxlWorkbook *wb, const char *name);
OxlWorksheet *oxl_workbook_get_sheet_by_index(OxlWorkbook *wb, uint32_t idx);
