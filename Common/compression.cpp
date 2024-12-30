#include "stdafx.h"
#include "compression.h"

#include "xcomdef.h"
#include "rtlcompression.h"

#pragma comment(lib, "ntdll.lib")

#ifndef STATUS_BUFFER_TOO_SMALL
#   define STATUS_BUFFER_TOO_SMALL 0xC0000023
#endif

#include <pshpack4.h>

struct COMPRESSION_HEADER
{
    WORD wFormat;
    DWORD dwUncompressedSize;
};

#include <poppack.h>

static buffer_t get_workspace(USHORT format)
{
    ULONG ulBufferSize, ulFragmentSize;
    auto status = RtlGetCompressionWorkSpaceSize(format, &ulBufferSize, &ulFragmentSize);
    if (IS_ERROR(status))
    {
        _com_raise_error(status);
    }
    return buffer_t(ulBufferSize);
}

buffer_t compress_buffer(const buffer_t& buf, USHORT format)
{
    buffer_t outbuf(sizeof(COMPRESSION_HEADER) + buf.size() + 0x1000);
    *(COMPRESSION_HEADER*)outbuf.data() = { format, (DWORD)buf.size() };

    ULONG outsize;
    format |= COMPRESSION_ENGINE_MAXIMUM;
    auto status = RtlCompressBuffer(
        format,
        (PUCHAR)buf.data(), (ULONG)buf.size(),
        outbuf.data() + sizeof(COMPRESSION_HEADER), (ULONG)outbuf.size() - sizeof(COMPRESSION_HEADER),
        4096,
        &outsize,
        get_workspace(format).data()
    );
    if (IS_ERROR(status))
    {
        _com_raise_error(status);
    }

    outbuf.resize(sizeof(COMPRESSION_HEADER) + outsize);
    return outbuf;
}

buffer_t decompress_buffer(const buffer_t& buf)
{
    if (buf.size() < sizeof(COMPRESSION_HEADER))
    {
        _com_raise_error(STATUS_BUFFER_TOO_SMALL);
    }

    ULONG outsize;
    auto header = (COMPRESSION_HEADER*)buf.data();
    buffer_t outbuf(header->dwUncompressedSize);
    auto status = RtlDecompressBufferEx(
        header->wFormat,
        outbuf.data(), (ULONG)outbuf.size(),
        (PUCHAR)buf.data() + sizeof(COMPRESSION_HEADER), (ULONG)buf.size() - sizeof(COMPRESSION_HEADER),
        &outsize,
        get_workspace(header->wFormat).data()
    );
    if (IS_ERROR(status))
    {
        _com_raise_error(status);
    }

    outbuf.resize(outsize);
    return outbuf;
}

