#include "string_table.h"
#include <stdlib.h>
#include <string.h>

/* Use khash for the write-path reverse map */
#include "khash.h"

KHASH_MAP_INIT_STR(str2idx, uint32_t)

void oxl_sst_init(OxlStringTable *sst) {
    sst->entries  = NULL;
    sst->count    = 0;
    sst->capacity = 0;
    sst->hash     = NULL;
}

void oxl_sst_free(OxlStringTable *sst) {
    for (uint32_t i = 0; i < sst->count; i++) free(sst->entries[i]);
    free(sst->entries);
    if (sst->hash) kh_destroy(str2idx, (khash_t(str2idx) *)sst->hash);
    sst->entries  = NULL;
    sst->count    = 0;
    sst->capacity = 0;
    sst->hash     = NULL;
}

uint32_t oxl_sst_append(OxlStringTable *sst, const char *str, uint32_t len) {
    if (sst->count == sst->capacity) {
        uint32_t cap = sst->capacity ? sst->capacity * 2 : 64;
        char **p = realloc(sst->entries, cap * sizeof(char *));
        if (!p) return UINT32_MAX;
        sst->entries  = p;
        sst->capacity = cap;
    }
    char *copy = malloc(len + 1);
    if (!copy) return UINT32_MAX;
    memcpy(copy, str, len);
    copy[len] = '\0';
    uint32_t idx = sst->count;
    sst->entries[sst->count++] = copy;
    return idx;
}

uint32_t oxl_sst_intern(OxlStringTable *sst, const char *str) {
    if (!sst->hash) sst->hash = kh_init(str2idx);
    khash_t(str2idx) *h = (khash_t(str2idx) *)sst->hash;

    int ret;
    khint_t k = kh_put(str2idx, h, str, &ret);
    if (ret == 0) {
        /* already present */
        return kh_value(h, k);
    }
    /* new entry */
    uint32_t idx = oxl_sst_append(sst, str, (uint32_t)strlen(str));
    /* the key in khash must point to the owned copy */
    kh_key(h, k) = sst->entries[idx];
    kh_value(h, k) = idx;
    return idx;
}

const char *oxl_sst_get(const OxlStringTable *sst, uint32_t idx) {
    if (idx >= sst->count) return NULL;
    return sst->entries[idx];
}
