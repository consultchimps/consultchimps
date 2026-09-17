/* ConsultChimps shim: the decoder-only XPress9 surface for WebAssembly.
 *
 * This file is ours and is licensed Apache-2.0 with the rest of the repository.
 * It mirrors the upstream src/Xpress9Wrapper.c (MIT, Hugoberry/xpress9-python,
 * commit 2503e827f95187d65c67ed491b184bf852256db7) minus the encoder, so the
 * encoder translation units never enter the wasm module.
 *
 * Exported surface, all through EMSCRIPTEN_KEEPALIVE and pinned again in
 * packages/pbi/scripts/build-wasm.mjs through EXPORTED_FUNCTIONS:
 *   _x9_create      allocate a decoder plus one started session
 *   _x9_destroy     free that decoder
 *   _x9_decompress  decode one framed chunk into a caller-owned buffer
 *   _x9_last_error  the decoder's own error text, for diagnostics only
 *   _x9_malloc      allocate linear memory for the JS caller
 *   _x9_free        release it
 *
 * The allocator is exported through _x9_malloc and _x9_free rather than through
 * emscripten's own _malloc and _free, so the module's JS surface stays exactly
 * these six entry points and nothing else.
 */
#include <stdlib.h>
#include <string.h>
#include "xpress.h"
#include "xpress9.h"
#include <emscripten/emscripten.h>

typedef struct {
    XPRESS9_DECODER decoder;
    char err[256];
} X9CTX;

static void *XPRESS_CALL AllocCb(void *ctx, int n) { (void)ctx; return malloc((size_t)n); }
static void XPRESS_CALL FreeCb(void *ctx, void *p) { (void)ctx; free(p); }

EMSCRIPTEN_KEEPALIVE
X9CTX *x9_create(void) {
    XPRESS9_STATUS st;
    X9CTX *c = (X9CTX *)calloc(1, sizeof(X9CTX));
    if (!c) return 0;
    memset(&st, 0, sizeof(st));
    c->decoder = Xpress9DecoderCreate(&st, NULL, AllocCb, XPRESS9_WINDOW_SIZE_LOG2_MAX, 0);
    if (c->decoder == NULL || st.m_uStatus != Xpress9Status_OK) { free(c); return 0; }
    memset(&st, 0, sizeof(st));
    Xpress9DecoderStartSession(&st, c->decoder, 1);
    if (st.m_uStatus != Xpress9Status_OK) {
        Xpress9DecoderDestroy(&st, c->decoder, NULL, FreeCb);
        free(c);
        return 0;
    }
    return c;
}

EMSCRIPTEN_KEEPALIVE
void x9_destroy(X9CTX *c) {
    XPRESS9_STATUS st;
    if (!c) return;
    memset(&st, 0, sizeof(st));
    if (c->decoder) Xpress9DecoderDestroy(&st, c->decoder, NULL, FreeCb);
    free(c);
}

/* Decompress one framed chunk. Returns bytes written, 0 on failure. */
EMSCRIPTEN_KEEPALIVE
unsigned x9_decompress(X9CTX *c, unsigned char *src, int srcLen, unsigned char *dst, int dstCap) {
    XPRESS9_STATUS st;
    unsigned total = 0, remaining;
    if (!c || !c->decoder) return 0;
    memset(&st, 0, sizeof(st));
    Xpress9DecoderAttach(&st, c->decoder, src, (unsigned)srcLen);
    if (st.m_uStatus != Xpress9Status_OK) {
        strncpy(c->err, st.m_ErrorDescription, sizeof(c->err) - 1);
        return 0;
    }
    do {
        unsigned written = 0, consumed = 0;
        remaining = Xpress9DecoderFetchDecompressedData(
            &st, c->decoder, dst + total, (unsigned)dstCap - total, &written, &consumed);
        if (st.m_uStatus != Xpress9Status_OK) {
            strncpy(c->err, st.m_ErrorDescription, sizeof(c->err) - 1);
            total = 0;
            break;
        }
        if (written == 0) break;
        total += written;
    } while (remaining != 0);
    memset(&st, 0, sizeof(st));
    Xpress9DecoderDetach(&st, c->decoder, src, (unsigned)srcLen);
    return total;
}

EMSCRIPTEN_KEEPALIVE
const char *x9_last_error(X9CTX *c) { return c ? c->err : ""; }

EMSCRIPTEN_KEEPALIVE
void *x9_malloc(int n) { return malloc((size_t)n); }

EMSCRIPTEN_KEEPALIVE
void x9_free(void *p) { free(p); }
