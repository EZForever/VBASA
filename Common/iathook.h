#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

void** iat_get(HMODULE hModule, LPCSTR szModule, LPCSTR szProc);

bool iat_modify(void** ppvEntry, void* pvValue);

inline void** iat_get_ordinal(HMODULE hModule, LPCSTR szModule, WORD wOrdinal)
{
    return iat_get(hModule, szModule, (LPCSTR)(IMAGE_ORDINAL_FLAG | wOrdinal));
}

