#ifndef NTSECURETSL_INTERNAL_H
#define NTSECURETSL_INTERNAL_H

#include <windows.h>
#include <sspi.h>
#include <schannel.h>

typedef struct _TLS_CTX_INTERNAL
{
    BOOL handshakeComplete;
    BOOL haveContext;
    CredHandle hCred;
    CtxtHandle hCtx;
    SECURITY_STATUS lastStatus;
} TLS_CTX_INTERNAL;

// Opaque handle type used by the public API
typedef void* TLS_CTX;

int __stdcall TlsClientHandshakeStep(
    void* h,
    const unsigned char* inBuf,  int inLen,
    unsigned char*       outBuf, int outBufSize,
    int*                 outLen,
    int*                 handshakeDone
);


#endif

