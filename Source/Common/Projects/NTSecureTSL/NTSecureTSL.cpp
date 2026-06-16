#include "NTSecureTSL.h"
#include "Internal.h"

BOOL APIENTRY DllMain(HINSTANCE hModule,
                      DWORD  ul_reason_for_call,
                      LPVOID lpReserved)
{
    return TRUE;
}




extern "C" __declspec(dllexport)
void* __stdcall TlsInit(const char* serverName, int* pErr)
{

    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)HeapAlloc(
        GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(TLS_CTX_INTERNAL));

    if (!ctx) {
        if (pErr) *pErr = -1;
        return NULL;
    }

	SCHANNEL_CRED cred;
	ZeroMemory(&cred, sizeof(cred));

	cred.dwVersion = SCHANNEL_CRED_VERSION;
	cred.grbitEnabledProtocols =
		SP_PROT_TLS1_0_CLIENT |
		SP_PROT_TLS1_1_CLIENT |
		SP_PROT_TLS1_2_CLIENT;

	cred.dwFlags =
		SCH_CRED_NO_DEFAULT_CREDS |
		SCH_CRED_MANUAL_CRED_VALIDATION;

	//cred.grbitEnabledProtocols = SP_PROT_TLS1_2_CLIENT;
	//cred.dwFlags |= SCH_USE_STRONG_CRYPTO;

	TimeStamp tsExpiry;
	SECURITY_STATUS ss= AcquireCredentialsHandle(
		NULL,
		UNISP_NAME,
		SECPKG_CRED_OUTBOUND,
		NULL,
		&cred,
		NULL,
		NULL,
		&ctx->hCred,
		&tsExpiry
	);


    if (ss != SEC_E_OK) {
		SetLastError((int)ss);
        HeapFree(GetProcessHeap(), 0, ctx);
        if (pErr) *pErr = (int)ss;
        return NULL;
    }

    ctx->haveContext = FALSE;
    ctx->handshakeComplete = FALSE;
    ctx->lastStatus = ss;

    if (pErr) *pErr = 0;
    return (void*)ctx;
}


extern "C" __declspec(dllexport)
void __stdcall TlsClose(void* h)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx) return;

    // Send close_notify if context is valid
    if (ctx->haveContext) {
        SecBuffer bufs[1];
        SecBufferDesc desc;

        bufs[0].BufferType = SECBUFFER_TOKEN;
        bufs[0].pvBuffer   = NULL;
        bufs[0].cbBuffer   = 0;

        desc.ulVersion = SECBUFFER_VERSION;
        desc.cBuffers  = 1;
        desc.pBuffers  = bufs;

        // Ask SChannel to generate close_notify
        SECURITY_STATUS ss = ApplyControlToken(&ctx->hCtx, &desc);

        if (ss == SEC_E_OK) {
            DWORD attr = 0;
            TimeStamp ts;

            ss = InitializeSecurityContextA(
                &ctx->hCred,
                &ctx->hCtx,
                NULL,
                ISC_REQ_SEQUENCE_DETECT |
                ISC_REQ_REPLAY_DETECT   |
                ISC_REQ_CONFIDENTIALITY |
                ISC_REQ_STREAM |
                ISC_REQ_ALLOCATE_MEMORY,
                0,
                SECURITY_NATIVE_DREP,
                NULL,
                0,
                &ctx->hCtx,
                &desc,
                &attr,
                &ts
            );

            if (ss == SEC_E_OK && bufs[0].pvBuffer) {
                // Normally you'd send this close_notify to the server,
                // but VB6 side handles socket close anyway.
                FreeContextBuffer(bufs[0].pvBuffer);
            }
        }
    }

    // Free context + creds
    if (ctx->haveContext)
        DeleteSecurityContext(&ctx->hCtx);

    FreeCredentialsHandle(&ctx->hCred);

    // Free our struct
    HeapFree(GetProcessHeap(), 0, ctx);
}

/*
extern "C" __declspec(dllexport)
int __stdcall TlsGetCipherInfo(void* h, char* outBuf, int outSize)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) return -1;

    SecPkgContext_ConnectionInfo info;
    SECURITY_STATUS ss = QueryContextAttributesA(
        &ctx->hCtx,
        SECPKG_ATTR_CONNECTION_INFO,
        &info
    );

    if (ss != SEC_E_OK) return ss;

    // Format: "TLS1.2, AES256, 256-bit"
    _snprintf(
        outBuf,
        outSize,
        "TLS%d.%d, CipherSuite=0x%04X, KeyBits=%d",
        info.dwProtocol >> 8,
        info.dwProtocol & 0xFF,
        info.aiCipher,
        info.dwCipherStrength
    );

    return 0;
}
*/

extern "C" __declspec(dllexport)
int __stdcall TlsRenegotiate(void* h)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) return -1;

    SECURITY_STATUS ss = ApplyControlToken(&ctx->hCtx, NULL);
    if (ss != SEC_E_OK) return ss;

    ctx->handshakeComplete = FALSE;
    ctx->haveContext = TRUE;   // keep same context
    ctx->lastStatus = 0;

    return 0;
}

extern "C" __declspec(dllexport)
int __stdcall TlsHandshake(
    void* ctx,
    const unsigned char* inBuf, int inLen,
    unsigned char* outBuf, int outSize,
    int* pErr)
{
    int outLen = 0;
    int done = 0;
    int rc = TlsClientHandshakeStep(ctx, inBuf, inLen, outBuf, outSize, &outLen, &done);
    if (pErr) *pErr = rc;   // rc will be 0 on SEC_I_CONTINUE_NEEDED / SEC_E_OK
    return outLen;
}
extern "C" __declspec(dllexport)
int __stdcall TlsIsHandshakeComplete(void* ctx)
{
    TLS_CTX_INTERNAL* c = (TLS_CTX_INTERNAL*)ctx;
    if (!c) return 0;
    return c->handshakeComplete ? 1 : 0;
}
extern "C" __declspec(dllexport)
int __stdcall TlsSend(
    void* h,
    const unsigned char* plain, int plainLen,
    unsigned char* outBuf, int outSize,
    int* pErr)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) {
        if (pErr) *pErr = -1;
        return 0;
    }

    SecBuffer bufs[4];
    SecBufferDesc desc;
    desc.ulVersion = SECBUFFER_VERSION;
    desc.cBuffers = 4;
    desc.pBuffers = bufs;

    // Layout for EncryptMessage:
    // [0] = header
    // [1] = data
    // [2] = trailer
    // [3] = empty

    // Header
    bufs[0].BufferType = SECBUFFER_STREAM_HEADER;
    bufs[0].cbBuffer   = 5;   // TLS record header size
    bufs[0].pvBuffer   = outBuf;

    // Data
    bufs[1].BufferType = SECBUFFER_DATA;
    bufs[1].cbBuffer   = plainLen;
    bufs[1].pvBuffer   = outBuf + bufs[0].cbBuffer;

    memcpy(bufs[1].pvBuffer, plain, plainLen);

    // Trailer
    bufs[2].BufferType = SECBUFFER_STREAM_TRAILER;
    bufs[2].cbBuffer   = outSize - (bufs[0].cbBuffer + bufs[1].cbBuffer);
    bufs[2].pvBuffer   = (unsigned char*)bufs[1].pvBuffer + bufs[1].cbBuffer;

    // Empty
    bufs[3].BufferType = SECBUFFER_EMPTY;
    bufs[3].cbBuffer   = 0;
    bufs[3].pvBuffer   = NULL;

    SECURITY_STATUS ss = EncryptMessage(&ctx->hCtx, 0, &desc, 0);

    if (ss != SEC_E_OK) {
        if (pErr) *pErr = ss;
        return 0;
    }

    int total =
        bufs[0].cbBuffer +
        bufs[1].cbBuffer +
        bufs[2].cbBuffer;

    if (pErr) *pErr = 0;
    return total;
}


extern "C" __declspec(dllexport)
int __stdcall TlsRecv(
    void* h,
    const unsigned char* enc, int encLen,
    unsigned char* outBuf, int outSize,
    int* pErr)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) {
        if (pErr) *pErr = -1;
        return 0;
    }

    // Copy encrypted data into a mutable buffer
    unsigned char* tmp = (unsigned char*)HeapAlloc(GetProcessHeap(), 0, encLen);
    memcpy(tmp, enc, encLen);

    SecBuffer bufs[4];
    SecBufferDesc desc;
    desc.ulVersion = SECBUFFER_VERSION;
    desc.cBuffers = 4;
    desc.pBuffers = bufs;

    bufs[0].BufferType = SECBUFFER_DATA;
    bufs[0].pvBuffer   = tmp;
    bufs[0].cbBuffer   = encLen;

    bufs[1].BufferType = SECBUFFER_EMPTY;
    bufs[2].BufferType = SECBUFFER_EMPTY;
    bufs[3].BufferType = SECBUFFER_EMPTY;

    SECURITY_STATUS ss = DecryptMessage(&ctx->hCtx, &desc, 0, NULL);

    if (ss == SEC_E_INCOMPLETE_MESSAGE) {
        if (pErr) *pErr = ss;
        HeapFree(GetProcessHeap(), 0, tmp);
        return 0;
    }

    if (ss != SEC_E_OK && ss != SEC_I_RENEGOTIATE) {
        if (pErr) *pErr = ss;
        HeapFree(GetProcessHeap(), 0, tmp);
        return 0;
    }

    // Find the decrypted data buffer
    int outLen = 0;
    for (int i = 0; i < 4; i++) {
        if (bufs[i].BufferType == SECBUFFER_DATA) {
            outLen = bufs[i].cbBuffer;
            if (outLen > outSize) outLen = outSize;
            memcpy(outBuf, bufs[i].pvBuffer, outLen);
            break;
        }
    }

    HeapFree(GetProcessHeap(), 0, tmp);

    if (pErr) *pErr = 0;
    return outLen;
}


extern "C" __declspec(dllexport)
int __stdcall TlsGetPeerCert(
    void* h,
    unsigned char* outBuf,
    int outSize,
    int* pErr)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) {
        if (pErr) *pErr = -1;
        return 0;
    }

    PCCERT_CONTEXT pCert = NULL;
    SECURITY_STATUS ss = QueryContextAttributesA(
        &ctx->hCtx,
        SECPKG_ATTR_REMOTE_CERT_CONTEXT,
        &pCert
    );

    if (ss != SEC_E_OK) {
        if (pErr) *pErr = ss;
        return 0;
    }

    int len = pCert->cbCertEncoded;
    if (len > outSize) len = outSize;

    memcpy(outBuf, pCert->pbCertEncoded, len);

    CertFreeCertificateContext(pCert);

    if (pErr) *pErr = 0;
    return len;
}

extern "C" __declspec(dllexport)
int __stdcall TlsValidateCert(
    void* h,
    const char* serverName,
    int* pErr)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) {
        if (pErr) *pErr = -1;
        return -1;
    }

    PCCERT_CONTEXT pCert = NULL;
    SECURITY_STATUS ss = QueryContextAttributesA(
        &ctx->hCtx,
        SECPKG_ATTR_REMOTE_CERT_CONTEXT,
        &pCert
    );

    if (ss != SEC_E_OK) {
        if (pErr) *pErr = ss;
        return -1;
    }

    HTTPSPolicyCallbackData polHttps;
    ZeroMemory(&polHttps, sizeof(polHttps));
    polHttps.cbStruct = sizeof(polHttps);
    polHttps.dwAuthType = AUTHTYPE_SERVER;
    polHttps.fdwChecks = 0;
    polHttps.pwszServerName = NULL;

    WCHAR wServer[256];
    MultiByteToWideChar(CP_ACP, 0, serverName, -1, wServer, 256);
    polHttps.pwszServerName = wServer;

    CERT_CHAIN_POLICY_PARA policyPara;
    ZeroMemory(&policyPara, sizeof(policyPara));
    policyPara.cbSize = sizeof(policyPara);
    policyPara.pvExtraPolicyPara = &polHttps;

    CERT_CHAIN_POLICY_STATUS policyStatus;
    ZeroMemory(&policyStatus, sizeof(policyStatus));
    policyStatus.cbSize = sizeof(policyStatus);

    CERT_CHAIN_PARA chainPara;
    ZeroMemory(&chainPara, sizeof(chainPara));
    chainPara.cbSize = sizeof(chainPara);

    PCCERT_CHAIN_CONTEXT pChain = NULL;

    if (!CertGetCertificateChain(
        NULL,
        pCert,
        NULL,
        pCert->hCertStore,
        &chainPara,
        0,
        NULL,
        &pChain))
    {
        CertFreeCertificateContext(pCert);
        if (pErr) *pErr = -2;
        return -1;
    }

    BOOL ok = CertVerifyCertificateChainPolicy(
        CERT_CHAIN_POLICY_SSL,
        pChain,
        &policyPara,
        &policyStatus
    );

    CertFreeCertificateChain(pChain);
    CertFreeCertificateContext(pCert);

    if (!ok || policyStatus.dwError != 0) {
        if (pErr) *pErr = policyStatus.dwError;
        return 0;   // invalid
    }

    if (pErr) *pErr = 0;
    return 1;       // valid
}

extern "C" __declspec(dllexport)
int __stdcall TlsGetProtocolVersion(void* h)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete)
        return -1;

    SecPkgContext_ConnectionInfo info;
    SECURITY_STATUS ss = QueryContextAttributesA(
        &ctx->hCtx,
        SECPKG_ATTR_CONNECTION_INFO,
        &info
    );

    if (ss != SEC_E_OK)
        return -1;

    return info.dwProtocol;   // e.g. SP_PROT_TLS1_2
}

/*
extern "C" __declspec(dllexport)
int __stdcall TlsGetPeerCertInfo(
    void* h,
    char* outBuf,
    int outSize,
    int* pErr)
{
    TLS_CTX_INTERNAL* ctx = (TLS_CTX_INTERNAL*)h;
    if (!ctx || !ctx->handshakeComplete) {
        if (pErr) *pErr = -1;
        return 0;
    }

    PCCERT_CONTEXT pCert = NULL;
    SECURITY_STATUS ss = QueryContextAttributesA(
        &ctx->hCtx,
        SECPKG_ATTR_REMOTE_CERT_CONTEXT,
        &pCert
    );

    if (ss != SEC_E_OK) {
        if (pErr) *pErr = ss;
        return 0;
    }

    char subject[512];
    char issuer[512];
    char validFrom[64];
    char validTo[64];

    CertNameToStrA(
        X509_ASN_ENCODING,
        &pCert->pCertInfo->Subject,
        CERT_X500_NAME_STR,
        subject,
        sizeof(subject)
    );

    CertNameToStrA(
        X509_ASN_ENCODING,
        &pCert->pCertInfo->Issuer,
        CERT_X500_NAME_STR,
        issuer,
        sizeof(issuer)
    );

    SYSTEMTIME stFrom, stTo;
    FileTimeToSystemTime(&pCert->pCertInfo->NotBefore, &stFrom);
    FileTimeToSystemTime(&pCert->pCertInfo->NotAfter, &stTo);

    wsprintfA(validFrom, "%04d-%02d-%02d", stFrom.wYear, stFrom.wMonth, stFrom.wDay);
    wsprintfA(validTo,   "%04d-%02d-%02d", stTo.wYear,   stTo.wMonth,   stTo.wDay);

    _snprintf(
        outBuf,
        outSize,
        "Subject=%s | Issuer=%s | Valid=%s to %s",
        subject,
        issuer,
        validFrom,
        validTo
    );

    CertFreeCertificateContext(pCert);

    if (pErr) *pErr = 0;
    return 1;
}
*/