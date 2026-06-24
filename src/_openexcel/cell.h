#pragma once
#include <stdint.h>

typedef enum {
    OXL_CELL_EMPTY      = 0,
    OXL_CELL_FLOAT      = 1,
    OXL_CELL_STRING     = 2,
    OXL_CELL_BOOL       = 3,
    OXL_CELL_ERROR      = 4,
    OXL_CELL_DATE       = 5,
    OXL_CELL_INLINE_STR = 6,
} OxlCellType;

typedef struct {
    int16_t  year;
    uint8_t  month;
    uint8_t  day;
    uint8_t  hour;
    uint8_t  min;
    uint8_t  sec;
    uint32_t usec;
} OxlDate;

typedef struct {
    uint32_t     row;       /* 0-based */
    uint16_t     col;       /* 0-based */
    uint16_t     style_idx;
    OxlCellType  type;
    union {
        double   f;         /* OXL_CELL_FLOAT */
        uint32_t s_idx;     /* OXL_CELL_STRING: index into OxlStringTable */
        char    *s_inline;  /* OXL_CELL_INLINE_STR or OXL_CELL_ERROR: owned heap ptr */
        int      b;         /* OXL_CELL_BOOL */
        OxlDate  dt;        /* OXL_CELL_DATE */
    } v;
} OxlCell;

/* Convert Excel serial date (days since 1900-01-00, with 1900 leap-year bug) to OxlDate.
   Pass date1904=1 to use the 1904 date system. */
OxlDate oxl_serial_to_date(double serial, int date1904);

/* Convert OxlDate back to an Excel serial for writing. */
double  oxl_date_to_serial(OxlDate d, int date1904);

void oxl_cell_free(OxlCell *c);
