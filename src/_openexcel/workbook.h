#pragma once
#include "worksheet.h"
#include "string_table.h"
#include "styles.h"

typedef struct {
    char    *name;           /* defined name, e.g. "MyRange" (owned) */
    char    *value;          /* formula/reference, e.g. "Sheet1!$A$1:$C$10" (owned) */
    int      local_sheet_id; /* 0-based sheet index, or -1 for global */
    int      hidden;         /* 1 if hidden */
} OxlDefinedName;

typedef struct {
    OxlWorksheet   **sheets;
    uint32_t         sheet_count;
    uint32_t         sheet_capacity;
    OxlStringTable   sst;
    OxlStyles        styles;
    int              date1904;  /* 0 = 1900 epoch (default), 1 = 1904 epoch */
    OxlDefinedName  *defined_names;
    uint32_t         defined_name_count;
    uint32_t         defined_name_capacity;
} OxlWorkbook;

OxlWorkbook *oxl_workbook_new(void);
void         oxl_workbook_free(OxlWorkbook *wb);

/* Add a sheet; workbook takes ownership. Returns sheet index or -1 on error. */
int          oxl_workbook_add_sheet(OxlWorkbook *wb, OxlWorksheet *ws);

OxlWorksheet *oxl_workbook_get_sheet_by_name(OxlWorkbook *wb, const char *name);
OxlWorksheet *oxl_workbook_get_sheet_by_index(OxlWorkbook *wb, uint32_t idx);

/* Add or replace a defined name. Returns 0 on success, -1 on OOM. */
int oxl_workbook_add_defined_name(OxlWorkbook *wb, const char *name,
                                   const char *value, int local_sheet_id, int hidden);
