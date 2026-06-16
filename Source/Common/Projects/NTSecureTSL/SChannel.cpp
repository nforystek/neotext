#define SECURITY_WIN32

#include "NTSecureTSL.h"
#include "Internal.h"


int __stdcall TlsClientHandshakeStep(
    void* h,
    const unsigned char* inBuf,  int inLen,
    unsigned char*       outBuf, int outBufSize,
    int*                 outLen,
    int*                 handshakeDone
)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx) return -1;

    *outLen = 0;
    *handshakeDone = ctx->handshakeComplete;

    SecBuffer inSecBuf[2];
    SecBufferDesc inDesc;
    inDesc.ulVersion = SECBUFFER_VERSION;
    inDesc.cBuffers = 2;
    inDesc.pBuffers = inSecBuf;

    if (inLen > 0 && inBuf != NULL) {
        inSecBuf[0].BufferType = SECBUFFER_TOKEN;
        inSecBuf[0].pvBuffer   = (void*)inBuf;
        inSecBuf[0].cbBuffer   = inLen;
    } else {
        inSecBuf[0].BufferType = SECBUFFER_EMPTY;
        inSecBuf[0].pvBuffer   = NULL;
        inSecBuf[0].cbBuffer   = 0;
    }

    inSecBuf[1].BufferType = SECBUFFER_EMPTY;
    inSecBuf[1].pvBuffer   = NULL;
    inSecBuf[1].cbBuffer   = 0;

    SecBuffer outSecBuf[1];
    SecBufferDesc outDesc;
    outDesc.ulVersion = SECBUFFER_VERSION;
    outDesc.cBuffers = 1;
    outDesc.pBuffers = outSecBuf;

    // Let SChannel allocate the output buffer
    outSecBuf[0].BufferType = SECBUFFER_TOKEN;
    outSecBuf[0].pvBuffer   = NULL;
    outSecBuf[0].cbBuffer   = 0;

    DWORD ctxReq =
        ISC_REQ_SEQUENCE_DETECT |
        ISC_REQ_REPLAY_DETECT   |
        ISC_REQ_CONFIDENTIALITY |
        ISC_REQ_STREAM |
        ISC_REQ_ALLOCATE_MEMORY;   // <-- important

    DWORD ctxAttr = 0;
    TimeStamp tsExpiry;
    SECURITY_STATUS ss;

    if (!ctx->haveContext) {
        ss = InitializeSecurityContextA(
            &ctx->hCred,
            NULL,
            NULL,
            ctxReq,
            0,
            SECURITY_NATIVE_DREP,
            (inLen > 0 ? &inDesc : NULL),
            0,
            &ctx->hCtx,
            &outDesc,
            &ctxAttr,
            &tsExpiry
        );
        if (ss == SEC_E_OK || ss == SEC_I_CONTINUE_NEEDED)
            ctx->haveContext = TRUE;
    } else {
        ss = InitializeSecurityContextA(
            &ctx->hCred,
            &ctx->hCtx,
            NULL,
            ctxReq,
            0,
            SECURITY_NATIVE_DREP,
            &inDesc,
            0,
            &ctx->hCtx,
            &outDesc,
            &ctxAttr,
            &tsExpiry
        );
    }



	ctx->lastStatus = ss;

	if (ss == SEC_E_OK) {
		ctx->handshakeComplete = TRUE;
		*handshakeDone = TRUE;
	}

	if (ss == SEC_E_OK || ss == SEC_I_CONTINUE_NEEDED) {
		if (outSecBuf[0].pvBuffer && outSecBuf[0].cbBuffer > 0) {
			int len = (int)outSecBuf[0].cbBuffer;
			if (len > outBufSize) len = outBufSize;

			memcpy(outBuf, outSecBuf[0].pvBuffer, len);
			*outLen = len;

			FreeContextBuffer(outSecBuf[0].pvBuffer);
		}
		return 0;   // <-- IMPORTANT
	}

	return (int)ss; // real error only


}
