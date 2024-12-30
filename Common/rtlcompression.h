#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#ifdef __cplusplus
extern "C" {
#endif

#ifndef NTSTATUS
#   define __NTSTATUS_NOT_DEFINED
    typedef HRESULT NTSTATUS; // XXX: They are basically the same thing
#endif

#define NT_RTL_COMPRESSION_API __declspec(dllimport) NTSTATUS WINAPI

NT_RTL_COMPRESSION_API RtlGetCompressionWorkSpaceSize(
    /*[in]*/  USHORT CompressionFormatAndEngine,
    /*[out]*/ PULONG CompressBufferWorkSpaceSize,
    /*[out]*/ PULONG CompressFragmentWorkSpaceSize
);

NT_RTL_COMPRESSION_API RtlCompressBuffer(
    /*[in]*/  USHORT CompressionFormatAndEngine,
    /*[in]*/  PUCHAR UncompressedBuffer,
    /*[in]*/  ULONG  UncompressedBufferSize,
    /*[out]*/ PUCHAR CompressedBuffer,
    /*[in]*/  ULONG  CompressedBufferSize,
    /*[in]*/  ULONG  UncompressedChunkSize,
    /*[out]*/ PULONG FinalCompressedSize,
    /*[in]*/  PVOID  WorkSpace
);

NT_RTL_COMPRESSION_API RtlDecompressBuffer(
    /*[in]*/  USHORT CompressionFormat,
    /*[out]*/ PUCHAR UncompressedBuffer,
    /*[in]*/  ULONG  UncompressedBufferSize,
    /*[in]*/  PUCHAR CompressedBuffer,
    /*[in]*/  ULONG  CompressedBufferSize,
    /*[out]*/ PULONG FinalUncompressedSize
);

NT_RTL_COMPRESSION_API RtlDecompressBufferEx(
    /*[in]*/  USHORT CompressionFormat,
    /*[out]*/ PUCHAR UncompressedBuffer,
    /*[in]*/  ULONG  UncompressedBufferSize,
    /*[in]*/  PUCHAR CompressedBuffer,
    /*[in]*/  ULONG  CompressedBufferSize,
    /*[out]*/ PULONG FinalUncompressedSize,
    /*[in]*/  PVOID  WorkSpace
);

NT_RTL_COMPRESSION_API RtlDecompressBufferEx2(
    /*[in]*/           USHORT CompressionFormat,
    /*[out]*/          PUCHAR UncompressedBuffer,
    /*[in]*/           ULONG  UncompressedBufferSize,
    /*[in]*/           PUCHAR CompressedBuffer,
    /*[in]*/           ULONG  CompressedBufferSize,
    /*[in]*/           ULONG  UncompressedChunkSize,
    /*[out]*/          PULONG FinalUncompressedSize,
    /*[in, optional]*/ PVOID  WorkSpace
);

#ifdef __NTSTATUS_NOT_DEFINED
#   undef NTSTATUS
#   undef __NTSTATUS_NOT_DEFINED
#endif

#ifdef __cplusplus
}
#endif

