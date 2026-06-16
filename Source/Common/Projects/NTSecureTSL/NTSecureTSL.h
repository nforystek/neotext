#ifndef NTSECURETSL_H
#define NTSECURETSL_H

#define SECURITY_WIN32
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <wincrypt.h>
#include <schannel.h>
#include <sspi.h>



// VC6 / Server 2003 SDK does not define modern TLS protocol flags
#ifndef SP_PROT_TLS1_0_CLIENT
#define SP_PROT_TLS1_0_CLIENT   0x00000040
#endif

#ifndef SP_PROT_TLS1_1_CLIENT
#define SP_PROT_TLS1_1_CLIENT   0x00000100
#endif

#ifndef SP_PROT_TLS1_2_CLIENT
#define SP_PROT_TLS1_2_CLIENT   0x00000200
#endif



// Missing protocol flags for older SDKs
#ifndef SP_PROT_TLS1_1_SERVER
#define SP_PROT_TLS1_1_SERVER 0x00000100
#endif
#ifndef SP_PROT_TLS1_1_CLIENT
#define SP_PROT_TLS1_1_CLIENT 0x00000200
#endif
#ifndef SP_PROT_TLS1_2_SERVER
#define SP_PROT_TLS1_2_SERVER 0x00000400
#endif
#ifndef SP_PROT_TLS1_2_CLIENT
#define SP_PROT_TLS1_2_CLIENT 0x00000800
#endif


#ifdef __cplusplus
extern "C" {
#endif

__declspec(dllexport) void* __stdcall TlsInit(const char* serverName, int* pErr);

__declspec(dllexport) int __stdcall TlsHandshake(
    void* ctx,
    const unsigned char* inBuf, int inLen,
    unsigned char* outBuf, int outSize,
    int* pErr);

__declspec(dllexport) int __stdcall TlsSend(
    void* ctx,
    const unsigned char* plain, int plainLen,
    unsigned char* outBuf, int outSize,
    int* pErr);

__declspec(dllexport) int __stdcall TlsRecv(
    void* ctx,
    const unsigned char* enc, int encLen,
    unsigned char* outBuf, int outSize,
    int* pErr);

__declspec(dllexport) void __stdcall TlsClose(void* ctx);

#ifdef __cplusplus
}
#endif

#endif // TLSWRAPPER_H
