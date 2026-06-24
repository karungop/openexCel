#pragma once
#include <stddef.h>
#include "../styles.h"

/* Parse styles.xml and populate OxlStyles with date xf flags.
   Returns 0 on success, -1 on error. */
int oxl_parse_styles(const char *buf, size_t len, OxlStyles *styles);
