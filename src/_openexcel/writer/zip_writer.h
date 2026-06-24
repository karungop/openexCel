#pragma once
#include <stddef.h>


typedef struct OxlZipWriter OxlZipWriter;

OxlZipWriter *oxl_zipw_open(const char *path);
void          oxl_zipw_close(OxlZipWriter *zw);

/* Add an in-memory buffer as a ZIP entry. Returns 0 on success. */
int oxl_zipw_add(OxlZipWriter *zw, const char *entry_name,
                 const char *data, size_t len);
