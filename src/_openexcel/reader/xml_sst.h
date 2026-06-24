#pragma once
#include <stddef.h>
#include "../string_table.h"

/* Parse sharedStrings.xml content into sst.
   buf must be null-terminated. Returns 0 on success, -1 on parse error. */
int oxl_parse_sst(const char *buf, size_t len, OxlStringTable *sst);
