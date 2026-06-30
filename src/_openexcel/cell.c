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

/* Convert a Gregorian date to a Julian Day Number (a standard astronomical day
   count). We use this to compute differences from Excel's epoch in O(1). */
static long gregorian_to_jdn(int y, int m, int dom) {
    /* Knuth / Richards algorithm */
    int a = (14 - m) / 12;
    int yy = y + 4800 - a;
    int mm = m + 12 * a - 3;
    return dom + (153 * mm + 2) / 5 + 365L * yy + yy / 4 - yy / 100 + yy / 400 - 32045;
}

double oxl_date_to_serial(OxlDate d, int date1904) {
    long days;
    if (date1904) {
        /* Day 0 = 1904-01-01 */
        long epoch = gregorian_to_jdn(1904, 1, 1);
        long target = gregorian_to_jdn(d.year, d.month, d.day);
        days = target - epoch;
    } else {
        /* Day 1 = 1900-01-01; epoch anchor is 1899-12-30 (so JDN diff gives serial) */
        long epoch = gregorian_to_jdn(1899, 12, 31);
        long target = gregorian_to_jdn(d.year, d.month, d.day);
        days = target - epoch;
        /* Excel 1900 bug: dates >= 1900-03-01 are off by 1 due to phantom Feb 29 */
        if (days >= 60) days++;
    }

    double frac = d.hour / 24.0 + d.min / 1440.0 + d.sec / 86400.0;
    return (double)days + frac;
}

void oxl_cell_free(OxlCell *c) {
    if (!c) return;
    if (c->type == OXL_CELL_INLINE_STR || c->type == OXL_CELL_ERROR) {
        free(c->v.s_inline);
        c->v.s_inline = NULL;
    }
    free(c->formula);
    c->formula = NULL;
    free(c->hyperlink);
    c->hyperlink = NULL;
}
