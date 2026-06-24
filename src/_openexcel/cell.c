#include "cell.h"
#include <stdlib.h>
#include <string.h>

/* Excel's 1900 date system has a deliberate bug: it treats 1900 as a leap year.
   Day 1 = 1900-01-01, Day 60 = 1900-02-29 (doesn't exist), Day 61 = 1900-03-01. */

static int is_leap(int y) {
    return (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
}

static int days_in_month(int m, int y) {
    static const int dim[] = {31,28,31,30,31,30,31,31,30,31,30,31};
    if (m == 2 && is_leap(y)) return 29;
    return dim[m - 1];
}

OxlDate oxl_serial_to_date(double serial, int date1904) {
    OxlDate d = {0};
    if (serial < 0.0) return d;

    long day_part;
    double frac;

    if (date1904) {
        /* 1904 system: day 0 = 1904-01-01 */
        day_part = (long)serial;
        frac = serial - day_part;
        int year = 1904, month = 1, day = 1;
        while (day_part > 0) {
            int dim = days_in_month(month, year);
            if (day_part <= dim) { day = (int)day_part; break; }
            day_part -= dim;
            month++;
            if (month > 12) { month = 1; year++; }
        }
        d.year = (int16_t)year;
        d.month = (uint8_t)month;
        d.day = (uint8_t)day;
    } else {
        /* 1900 system: skip the phantom day 60 (1900-02-29) */
        if (serial >= 60.0) serial -= 1.0;
        day_part = (long)serial;
        frac = serial - (double)(long)serial;
        /* day 1 = 1899-12-31; so day_part days after that */
        int year = 1899, month = 12, day = 31;
        while (day_part > 0) {
            int dim = days_in_month(month, year);
            int remaining = dim - day;
            if (day_part <= remaining) { day += (int)day_part; day_part = 0; }
            else {
                day_part -= remaining + 1;
                day = 1;
                month++;
                if (month > 12) { month = 1; year++; }
            }
        }
        d.year = (int16_t)year;
        d.month = (uint8_t)month;
        d.day = (uint8_t)day;
    }

    /* time from fractional part */
    long total_sec = (long)(frac * 86400.0 + 0.5);
    d.hour = (uint8_t)(total_sec / 3600);
    d.min  = (uint8_t)((total_sec % 3600) / 60);
    d.sec  = (uint8_t)(total_sec % 60);
    return d;
}

double oxl_date_to_serial(OxlDate d, int date1904) {
    /* Count days from epoch to d.year-d.month-d.day */
    long days = 0;
    int base_year, base_month, base_day;

    if (date1904) {
        base_year = 1904; base_month = 1; base_day = 1;
    } else {
        base_year = 1899; base_month = 12; base_day = 31;
    }

    int y = base_year, m = base_month, day = base_day;
    while (y < d.year || (y == d.year && m < d.month) ||
           (y == d.year && m == d.month && day < d.day)) {
        days++;
        day++;
        if (day > days_in_month(m, y)) { day = 1; m++; }
        if (m > 12) { m = 1; y++; }
    }

    if (!date1904 && days >= 59) days++; /* re-insert the phantom day 60 */

    double frac = d.hour / 24.0 + d.min / 1440.0 + d.sec / 86400.0;
    return (double)days + frac;
}

void oxl_cell_free(OxlCell *c) {
    if (!c) return;
    if (c->type == OXL_CELL_INLINE_STR || c->type == OXL_CELL_ERROR) {
        free(c->v.s_inline);
        c->v.s_inline = NULL;
    }
}
