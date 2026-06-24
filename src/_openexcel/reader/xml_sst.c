#include "xml_sst.h"
#include <expat.h>
#include <stdlib.h>
#include <string.h>

typedef enum {
    SST_NONE,
    SST_SI,     /* inside <si> */
    SST_T,      /* inside <si><t> or <si><r><t> */
} SstState;

typedef struct {
    OxlStringTable *sst;
    SstState state;
    /* accumulate text across multiple charData callbacks */
    char    *cbuf;
    size_t   cbuf_len;
    size_t   cbuf_cap;
    int      error;
    /* For richText <si><r><t>: accumulate all <t> segments */
    char    *rich_buf;
    size_t   rich_len;
    size_t   rich_cap;
    int      in_rich; /* inside <si><r> */
} SstCtx;

static void cbuf_append(char **buf, size_t *len, size_t *cap, const char *s, int n) {
    size_t need = *len + (size_t)n;
    if (need >= *cap) {
        size_t newcap = *cap ? *cap * 2 : 256;
        while (newcap <= need) newcap *= 2;
        char *p = realloc(*buf, newcap);
        if (!p) return;
        *buf = p;
        *cap = newcap;
    }
    memcpy(*buf + *len, s, (size_t)n);
    *len += (size_t)n;
}

static void XMLCALL sst_start(void *ud, const char *name, const char **attrs) {
    (void)attrs;
    SstCtx *c = ud;
    /* Strip namespace prefix if present */
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "si") == 0) {
        c->state = SST_SI;
        c->rich_len = 0;
        c->in_rich = 0;
    } else if (strcmp(name, "r") == 0 && c->state == SST_SI) {
        c->in_rich = 1;
    } else if (strcmp(name, "t") == 0 && c->state == SST_SI) {
        c->state = SST_T;
        c->cbuf_len = 0;
    }
}

static void XMLCALL sst_end(void *ud, const char *name) {
    SstCtx *c = ud;
    const char *local = strrchr(name, ':');
    if (local) name = local + 1;

    if (strcmp(name, "t") == 0 && c->state == SST_T) {
        if (c->in_rich) {
            /* Append this run's text to rich_buf */
            cbuf_append(&c->rich_buf, &c->rich_len, &c->rich_cap,
                        c->cbuf, (int)c->cbuf_len);
        }
        c->state = SST_SI;
    } else if (strcmp(name, "r") == 0 && c->state == SST_SI) {
        c->in_rich = 0;
    } else if (strcmp(name, "si") == 0) {
        /* Commit this string to the table */
        const char *text;
        size_t tlen;
        if (c->rich_len > 0) {
            /* Rich text: use concatenated plain text */
            text = c->rich_buf;
            tlen = c->rich_len;
        } else {
            text = c->cbuf ? c->cbuf : "";
            tlen = c->cbuf_len;
        }
        oxl_sst_append(c->sst, text, (uint32_t)tlen);
        c->cbuf_len  = 0;
        c->rich_len  = 0;
        c->state = SST_NONE;
    }
}

static void XMLCALL sst_char(void *ud, const char *s, int n) {
    SstCtx *c = ud;
    if (c->state == SST_T) {
        cbuf_append(&c->cbuf, &c->cbuf_len, &c->cbuf_cap, s, n);
    }
}

int oxl_parse_sst(const char *buf, size_t len, OxlStringTable *sst) {
    SstCtx c = {0};
    c.sst = sst;

    XML_Parser p = XML_ParserCreate("UTF-8");
    if (!p) return -1;
    XML_SetUserData(p, &c);
    XML_SetElementHandler(p, sst_start, sst_end);
    XML_SetCharacterDataHandler(p, sst_char);

    int ok = XML_Parse(p, buf, (int)len, 1) == XML_STATUS_OK;
    XML_ParserFree(p);
    free(c.cbuf);
    free(c.rich_buf);
    return (ok && !c.error) ? 0 : -1;
}
