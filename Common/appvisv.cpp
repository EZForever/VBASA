#include "stdafx.h"
#include "appvisv.h"

#if defined(_M_IX86)
#   define APPVISV_MODULE "AppvIsvSubsystems32.dll"
#elif defined(_M_AMD64)
#   define APPVISV_MODULE "AppvIsvSubsystems64.dll"
#else
#   error "Unsupported platform"
#endif

typedef void (WINAPI* VirtualizeCurrentProcess_t)(BOOL bEnable);

static HMODULE g_hModule;

bool init_appvisv()
{
	if (g_hModule)
	{
		return true;
	}

    // ---

    wchar_t szModulePath[MAX_PATH];
    DWORD cbModulePath = sizeof(szModulePath);
    auto ret = RegGetValueW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Office\\ClickToRun", L"PackageFolder", RRF_RT_REG_SZ, NULL, szModulePath, &cbModulePath);
    if (ret)
    {
        SetLastError(ret);
        return false;
    }
    wcscat_s(szModulePath, L"\\root\\Client\\" APPVISV_MODULE);

    wchar_t szModuleFullPath[MAX_PATH];
    if (!GetFullPathNameW(szModulePath, _countof(szModuleFullPath), szModuleFullPath, NULL))
    {
        return false;
    }

    // ---

    g_hModule = LoadLibraryExW(szModuleFullPath, NULL, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    if (!g_hModule)
    {
        return false;
    }
    atexit([]() { FreeLibrary(g_hModule); g_hModule = NULL; });

    auto VirtualizeCurrentProcess = (VirtualizeCurrentProcess_t)GetProcAddress(g_hModule, "VirtualizeCurrentProcess");
    if (!VirtualizeCurrentProcess)
    {
        return false;
    }

    VirtualizeCurrentProcess(TRUE);
    return true;
}

