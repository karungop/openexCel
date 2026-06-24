#pragma once
#include <stdint.h>

typedef struct {
    char    **entries;    /* entries[i] = owned UTF-8 string */
    uint32_t  count;
    uint32_t  capacity;
    void     *hash;       /* khash str→idx map, only populated for write path */
} OxlStringTable;

void     oxl_sst_init(OxlStringTable *sst);
void     oxl_sst_free(OxlStringTable *sst);

/* Append a new string (copied). Returns its index. */
uint32_t oxl_sst_append(OxlStringTable *sst, const char *str, uint32_t len);

/* Intern a string for writing: returns existing index or appends. O(1) average. */
uint32_t oxl_sst_intern(OxlStringTable *sst, const char *str);

/* Return string at index, or NULL if out of range. */
const char *oxl_sst_get(const OxlStringTable *sst, uint32_t idx);
