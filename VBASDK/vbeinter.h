#pragma once

#include "vbe.h"

// Remove this to revert to original, VBA6 behavior
#define VBA7

// Load Debug build of VBA instead of Retail build
//#define VBADEBUG

// Enable VBASA specific features
#define VBASA

// ---

namespace VBE
{
#ifdef VBA7
	typedef LPCSTR VBACOOKIE;
#else
	typedef DWORD VBACOOKIE;
#endif

	// Alternate value of szCookie, used internally by Microsoft Office
	// Causes subtle behavior changes in VBE
	const auto VBACOOKIE_OFFICE = (VBACOOKIE)0x0FF1CE;
}

// ---

#ifdef VBASA

// VBASA specific: Implement this function to receive notification on VBA DLL load
extern "C" bool _VbaLoadHook();

#endif

// ---

// Called during host application startup to initialize VBA.
// The pVbaInitInfo parameter points to a VBAINITINFO structure, which in turn contains a VBAAPPOBJDESC member named vaodHost. Initialize the uVerMaj and uVerMin members of vaodHost to the current major and minor versions (respectively) of the type library file.
STDAPI VbaInitialize(VBE::VBACOOKIE dwCookie, VBE::VBAINITINFO* pVbaInitInfo, VBE::IVba** ppVba);

// Called by the host if it wants to support add-ins. Writes a pointer to the interface to ppHostAddins.
// *Note* This function should be called only once per session.
STDAPI VbaInitializeHostAddins(VBE::VBAINITHOSTADDININFO* pInitAddinInfo, VBE::IVbaHostAddins** ppHostAddins);

// Called just before the application closes. All projects should already be closed.
STDAPI VbaTerminate(VBE::IVba* pVba);

// szCookie must be VBECOOKIE_OFFICE, otherwise the call will fail with error E_ACCESSDENIED
// Set pVbaHash->vbahashtype before calling
STDAPI VbeGetHashOfCode(VBE::VBACOOKIE dwCookie, IStorage* pStg, VBE::VBAHASH* pVbaHash);

// ???
STDAPI VbeSetGimmeTable();

