#pragma once
#include <type_traits>

#include "vbe.h"
#include "vbeinter.h"

struct VbaHelperParams
{
	VBE::VBACOOKIE cookie;
	LCID lcid = LOCALE_USER_DEFAULT;
	UINT nameid;
	bool headless;
	bool mdehack = false;
};

class VbaHelper final
{
private:
	static VbaHelper g_Instance;

	VBE::IVbaPtr m_vba;
	VBE::IVbaSitePtr m_vbaSite;
	VBE::IVbaCompManagerSitePtr m_vbaCompManagerSite;

	inline VbaHelper() = default;
	inline VbaHelper(VbaHelper&& rhs) noexcept { *this = std::move(rhs); }
	VbaHelper(const VbaHelperParams& params);
	~VbaHelper();

	VbaHelper& operator=(VbaHelper&& rhs) noexcept;

public:
	static inline void CreateInstance(const VbaHelperParams& params)
		{ g_Instance = VbaHelper(params); }
	static inline VbaHelper& GetInstance()
		{ return g_Instance; }

	inline VBE::IVbaPtr GetVba() { return m_vba; }

	void Reset();
	long MessageLoop();
	VBE::IVbaProjectPtr LoadProject(IStoragePtr pStg, bool initnew = false);
};

