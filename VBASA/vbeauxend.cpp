#include "stdafx.h"
#include "vbeauxend.h"

#include "xcomdef.h"

#include "VbaHelper.h"

// HACK: Hardcoded structure offsets
// XXX: x86 offsets not tested
#if defined(_M_IX86)
#	define OFF_VBAPROJ_TO_PROJECT 15
#	define OFF_PROJECT_TO_GENPROJ 42
#	define OFF_VBAPROJINT_VTBL_SETOPTIONS 35
#elif defined(_M_AMD64)
#	define OFF_VBAPROJ_TO_PROJECT 14
#	define OFF_PROJECT_TO_GENPROJ 36
#	define OFF_VBAPROJINT_VTBL_SETOPTIONS 35
#else
#	error "Current impl uses x64 offsets"
#endif

// ---

#include <initguid.h>

// VBE7!IID_IVbaProject
// NOTE: Not the same interface as VBE::IVbaProject
DEFINE_GUID(IID_IVbaProjectInternal, 0xDDD557E0, 0xD96F, 0x11CD, 0x95, 0x70, 0x00, 0xaa, 0x00, 0x51, 0xE5, 0xD4);

// ---

// VBE7!GEN_PROJECT::SetOptions
typedef HRESULT (STDMETHODCALLTYPE* VbaProjectInternal_SetOptions_t)(void* pThis, ULONG ulOptions);

static VbaProjectInternal_SetOptions_t g_pVbaProjIntSetOptions;

static STDMETHODIMP xVbaProjectInternal_SetOptions(void* pThis, ULONG ulOptions)
{
	return g_pVbaProjIntSetOptions(pThis, ulOptions & ~1);
}

// ---

bool init_vbeaux_end()
{
	bool ret = true;
	VBE::IVbaProjectPtr vbaProj;
	try
	{
		HRESULT hr;

		ILockBytesPtr lkbyt;
		_XCOM_RAISE_IF_FAILED(CreateILockBytesOnHGlobal(NULL, TRUE, &lkbyt));

		IStoragePtr stg;
		_XCOM_RAISE_IF_FAILED(StgCreateDocfileOnILockBytes(lkbyt, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, &stg));

		vbaProj = VbaHelper::GetInstance().LoadProject(stg, true);

		auto pProject = *((void**)vbaProj.GetInterfacePtr() + OFF_VBAPROJ_TO_PROJECT); // VBE7!CVbeProject::InitNewProject
		auto pGenProj = *((void**)pProject + OFF_PROJECT_TO_GENPROJ); // VBE7!Project::OpenStg

		IUnknownPtr vbaProjInt;
		IUnknownPtr((IUnknown*)pGenProj, true).QueryInterface(IID_IVbaProjectInternal, &vbaProjInt);

		auto pVbaProjIntVtbl = *(void**)vbaProjInt.GetInterfacePtr(); // vtbl
		auto ppVbaProjIntSetOptions = (VbaProjectInternal_SetOptions_t*)pVbaProjIntVtbl + OFF_VBAPROJINT_VTBL_SETOPTIONS; // VBE7!GEN_PROJECT::'vftable'{for 'IVbaProject'}
		
		DWORD flOldProtect;
		BOOL fRet = VirtualProtect(ppVbaProjIntSetOptions, sizeof(*ppVbaProjIntSetOptions), PAGE_READWRITE, &flOldProtect);
		if (fRet)
		{
			g_pVbaProjIntSetOptions = *ppVbaProjIntSetOptions;
			*ppVbaProjIntSetOptions = xVbaProjectInternal_SetOptions;
			VirtualProtect(ppVbaProjIntSetOptions, sizeof(*ppVbaProjIntSetOptions), flOldProtect, &flOldProtect);
		}
		else
		{
			ret = false;
		}
	}
	catch (...)
	{
		ret = false;
	}
	if (vbaProj)
	{
		vbaProj->Close();
	}
	return ret;
}

