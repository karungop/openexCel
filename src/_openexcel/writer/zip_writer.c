#include "zip_writer.h"
#include "miniz.h"
#include <stdlib.h>
#include <string.h>

struct OxlZipWriter {
    mz_zip_archive zip;
    int ok;
};

OxlZipWriter *oxl_zipw_open(const char *path) {
    OxlZipWriter *zw = calloc(1, sizeof(OxlZipWriter));
    if (!zw) return NULL;
    if (!mz_zip_writer_init_file(&zw->zip, path, 0)) {
        free(zw);
        return NULL;
    }
    zw->ok = 1;
    return zw;
}

void oxl_zipw_close(OxlZipWriter *zw) {
    if (!zw) return;
    if (zw->ok) mz_zip_writer_finalize_archive(&zw->zip);
    mz_zip_writer_end(&zw->zip);
    free(zw);
}

int oxl_zipw_add(OxlZipWriter *zw, const char *entry_name,
                 const char *data, size_t len) {
    if (!zw->ok) return -1;
    if (!mz_zip_writer_add_mem(&zw->zip, entry_name, data, len,
                               MZ_BEST_SPEED)) {
        zw->ok = 0;
        return -1;
    }
    return 0;
}
