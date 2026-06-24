#include "zip_reader.h"
#include "miniz.h"
#include <stdlib.h>
#include <string.h>

struct OxlZipReader {
    mz_zip_archive zip;
};

OxlZipReader *oxl_zip_open(const char *path) {
    OxlZipReader *zr = calloc(1, sizeof(OxlZipReader));
    if (!zr) return NULL;
    if (!mz_zip_reader_init_file(&zr->zip, path, 0)) {
        free(zr);
        return NULL;
    }
    return zr;
}

void oxl_zip_close(OxlZipReader *zr) {
    if (!zr) return;
    mz_zip_reader_end(&zr->zip);
    free(zr);
}

int oxl_zip_extract(OxlZipReader *zr, const char *entry_name,
                    char **out_buf, size_t *out_size) {
    int idx = mz_zip_reader_locate_file(&zr->zip, entry_name, NULL, 0);
    if (idx < 0) return -1;

    mz_zip_archive_file_stat stat;
    if (!mz_zip_reader_file_stat(&zr->zip, (mz_uint)idx, &stat)) return -1;

    size_t sz = (size_t)stat.m_uncomp_size;
    char *buf = malloc(sz + 1);
    if (!buf) return -1;

    if (!mz_zip_reader_extract_to_mem(&zr->zip, (mz_uint)idx, buf, sz, 0)) {
        free(buf);
        return -1;
    }
    buf[sz] = '\0';
    *out_buf  = buf;
    *out_size = sz;
    return 0;
}

uint32_t oxl_zip_file_count(OxlZipReader *zr) {
    return (uint32_t)mz_zip_reader_get_num_files(&zr->zip);
}

const char *oxl_zip_filename(OxlZipReader *zr, uint32_t i) {
    static char namebuf[512];
    if (!mz_zip_reader_get_filename(&zr->zip, i, namebuf, sizeof(namebuf))) return NULL;
    return namebuf;
}
