#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <vector>

typedef std::vector<BYTE> buffer_t;

buffer_t compress_buffer(const buffer_t& buf, USHORT format = COMPRESSION_FORMAT_XPRESS_HUFF);

buffer_t decompress_buffer(const buffer_t& buf);

