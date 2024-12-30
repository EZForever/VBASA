#include "stdafx.h"
#include "resource.h"
#include "build.h"

#include "xcomdef.h"
#include "compression.h"

#include "VbaHelper.h"

static const WCHAR g_wszRunnerName[] = L"VBASARun.exe";

static inline LPCWSTR ids_to_name(WORD wResId)
{
    // https://stackoverflow.com/a/14089163
    return MAKEINTRESOURCEW((wResId >> 4) + 1);
}

static LPVOID get_resource(LPCWSTR lpName, LPCWSTR lpType, DWORD* pcbData)
{
    *pcbData = 0;

    HRSRC hrsrc = FindResourceW(NULL, lpName, lpType);
    if (!hrsrc)
    {
        return NULL;
    }

    HGLOBAL hResData = LoadResource(NULL, hrsrc);
    if (!hResData)
    {
        return NULL;
    }

    *pcbData = SizeofResource(NULL, hrsrc);
    return LockResource(hResData);
}

void build_project(_bstr_t filename, _bstr_t buildname, bool strip)
{
    HRESULT hr;

    // The project file is opened by CProjSite::Create with STGM_SHARE_EXCLUSIVE, thus a temp copy would be necessary
    WCHAR tempname[MAX_PATH];
    _XCOM_RAISE_IF_WIN32(GetTempPathW(_countof(tempname), tempname) == 0);
    _XCOM_RAISE_IF_WIN32(GetTempFileNameW(tempname, L"VBA", 0, tempname) == 0);
    _XCOM_RAISE_IF_WIN32(!CopyFileW(filename, tempname, FALSE));

    ILockBytesPtr lkbyt;
    try
    {
        IStoragePtr stgIn;
        _XCOM_RAISE_IF_FAILED(StgOpenStorage(tempname, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &stgIn));

        auto proj = VbaHelper::GetInstance().LoadProject(stgIn);
        try
        {
            _XCOM_RAISE_IF_FAILED(CreateILockBytesOnHGlobal(NULL, TRUE, &lkbyt));

            IStoragePtr stgOut;
            _XCOM_RAISE_IF_FAILED(StgCreateDocfileOnILockBytes(lkbyt, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, &stgOut));
            if (strip)
            {
                VBE::IVbaProjectExPtr(proj)->MakeMDE(stgOut);
            }
            else
            {
                IPersistStoragePtr(proj)->Save(stgOut, FALSE);
            }
        }
        catch (...)
        {
            proj->Close();
            throw;
        }
        proj->Close();
    }
    catch (...)
    {
        DeleteFileW(tempname);
        throw;
    }
    DeleteFileW(tempname);

    // ---

    STATSTG statstg;
    lkbyt->Stat(&statstg, STATFLAG_NONAME);
    if (statstg.cbSize.HighPart != 0)
    {
        // C'mon, a VBA project isn't gonna be 2GB+ in size
        _com_raise_error(E_OUTOFMEMORY);
    }

    ULONG cb = statstg.cbSize.LowPart;
    buffer_t buf(cb);
    lkbyt->ReadAt(ULARGE_INTEGER{ 0 }, buf.data(), cb, &cb);
    
    buf = compress_buffer(buf);

    // ---

    _XCOM_RAISE_IF_WIN32(SearchPathW(NULL, L"VBASARun", L".exe", _countof(tempname), tempname, NULL) == 0);
    _XCOM_RAISE_IF_WIN32(!CopyFileW(tempname, buildname, FALSE));
    try
    {
        auto hUpdate = BeginUpdateResourceW(buildname, FALSE);
        _XCOM_RAISE_IF_WIN32(!hUpdate);
        try
        {
            _XCOM_RAISE_IF_WIN32(!UpdateResourceW(hUpdate, RT_RCDATA, L"VBAPROJ", MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), buf.data(), (DWORD)buf.size()));
            
            // Replace string table IDS_APPNAME -> IDS_BLDAPPNAME
            DWORD cbData;
            auto pData = get_resource(ids_to_name(IDS_BLDAPPNAME), RT_STRING, &cbData);
            _XCOM_RAISE_IF_WIN32(!UpdateResourceW(hUpdate, RT_STRING, ids_to_name(IDS_APPNAME), MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), pData, cbData));

            // Replace version info VS_VERSION_INFO -> IDR_BLDVERSION
            pData = get_resource(MAKEINTRESOURCEW(IDR_BLDVERSION), RT_VERSION, &cbData);
            _XCOM_RAISE_IF_WIN32(!UpdateResourceW(hUpdate, RT_VERSION, MAKEINTRESOURCEW(VS_VERSION_INFO), MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), pData, cbData));
        }
        catch (...)
        {
            EndUpdateResourceW(hUpdate, TRUE);
            throw;
        }
        EndUpdateResourceW(hUpdate, FALSE);
    }
    catch(...)
    {
        DeleteFileW(buildname);
        throw;
    }
}

