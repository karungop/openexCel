#pragma once
#include <stddef.h>
#include <stdint.h>

/* Thin wrapper around miniz for reading ZIP (xlsx) archives. */

typedef struct OxlZipReader OxlZipReader;

OxlZipReader *oxl_zip_open(const char *path);
void          oxl_zip_close(OxlZipReader *zr);

/* Extract a ZIP entry by name into a newly malloc'd buffer.
   Caller must free(*out_buf). Returns 0 on success, -1 if not found or error. */
int oxl_zip_extract(OxlZipReader *zr, const char *entry_name,
                    char **out_buf, size_t *out_size);

/* Returns number of files in the archive. */
uint32_t oxl_zip_file_count(OxlZipReader *zr);

/* Returns the name of file at index i, or NULL. */
const char *oxl_zip_filename(OxlZipReader *zr, uint32_t i);
