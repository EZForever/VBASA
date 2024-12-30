#include "stdafx.h"
#include "iathook.h"

#pragma comment(lib, "dbghelp.lib")

static LONG g_lIatLock;

void** iat_get(HMODULE hModule, LPCSTR szModule, LPCSTR szProc)
{
    auto pNtHeaders = ImageNtHeader(hModule);
    if (!pNtHeaders)
    {
        return nullptr;
    }

    ULONG cbImportDir;
    auto pImportDir = (PIMAGE_IMPORT_DESCRIPTOR)ImageDirectoryEntryToDataEx(hModule, TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, &cbImportDir, NULL);
    if (!pImportDir)
    {
        return nullptr;
    }

    PIMAGE_THUNK_DATA pLookupTable = NULL;
    PIMAGE_THUNK_DATA pAddressTable = NULL;
    for (int i = 0; i < cbImportDir / sizeof(IMAGE_IMPORT_DESCRIPTOR); i++)
    {
        auto& entry = pImportDir[i];
        if (!entry.Name || !entry.OriginalFirstThunk || !entry.FirstThunk)
        {
            continue;
        }

        auto pName = (const char*)((PBYTE)hModule + entry.Name);
        if (pName && !_stricmp(pName, szModule))
        {
            pLookupTable = (PIMAGE_THUNK_DATA)((PBYTE)hModule + entry.OriginalFirstThunk);
            pAddressTable = (PIMAGE_THUNK_DATA)((PBYTE)hModule + entry.FirstThunk);
            break;
        }
    }
    if (!pLookupTable || !pAddressTable)
    {
        return nullptr;
    }

    WORD wProcOrdinal = IMAGE_SNAP_BY_ORDINAL((UINT_PTR)szProc) ? IMAGE_ORDINAL((UINT_PTR)szProc) : 0;
    while (pLookupTable->u1.Ordinal)
    {
        WORD wOrdinal = IMAGE_SNAP_BY_ORDINAL(pLookupTable->u1.Ordinal) ? IMAGE_ORDINAL(pLookupTable->u1.Ordinal) : 0;
        if (wOrdinal == wProcOrdinal)
        {
            auto pNameTable = (PIMAGE_IMPORT_BY_NAME)((PBYTE)hModule + pLookupTable->u1.AddressOfData);
            if (wOrdinal != 0 || !strcmp(pNameTable->Name, szProc))
            {
                return (void**)&pAddressTable->u1.Function;
            }
        }

        pLookupTable++;
        pAddressTable++;
    }

    return nullptr;
}

bool iat_modify(void** ppvEntry, void* pvValue)
{
    // Yes, this is a spinlock; iat_modify shouldn't take that long though
    while (InterlockedCompareExchange(&g_lIatLock, -1, 0) != 0)
    {
        YieldProcessor();
    }

    DWORD flOldProtect;
    BOOL fRet = VirtualProtect(ppvEntry, sizeof(*ppvEntry), PAGE_READWRITE, &flOldProtect);
    if (fRet)
    {
        *ppvEntry = pvValue;
        VirtualProtect(ppvEntry, sizeof(*ppvEntry), flOldProtect, &flOldProtect);
    }

    InterlockedAnd(&g_lIatLock, 0);
    return fRet;
}

