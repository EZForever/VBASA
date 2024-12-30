#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#define GDIPVER 0x0110
#include <objidl.h>
#include <gdiplus.h>
#include <mmsystem.h>

#define _USE_MATH_DEFINES
#include <cmath>
#include <array>
#include <tuple>
#include <memory>
#include <algorithm>
#include <initializer_list>

#include "resource.h"
#include "ui.h"

#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "gdiplus.lib")

#pragma warning(disable: 4244) // Loss of accurancy casting float from/to int

#define ENABLE_MIDI

using point3_t = std::array<float, 3>; // 3d coord
using point2_t = Gdiplus::PointF;
using point2d_t = std::tuple<point2_t, float>; // 2d coord w/ depth

using face3_t = std::array<point3_t, 4>;
using face2_t = std::array<point2_t, 4>;
using face2d_t = std::tuple<face2_t, float>;

using cube3_t = std::array<face3_t, 6>;
using cube2_t = std::array<face2_t, 6>;
using cube2d_t = std::tuple<cube2_t, float>;

static struct
{
	BOOL fDrawnOnce;

	ATOM atomOfVbEg;
	WCHAR wszText[512];
	WCHAR wszTitle[64];
	LPVOID pvMidi;
	DWORD cbMidi;
	LPBITMAPINFO pbmiBackground;
	DWORD cbBackground;

	// ---

	HWND hWnd;
	ULONG_PTR ulpGdiplusToken;

	WCHAR wszMidiFileName[MAX_PATH];
	MCIDEVICEID wMidiDevice;

	// ---

	int iWndWidth;
	int iWndHeight;

	int iTicks;
	float fBgSpeed;
	float fCubeRotSpeed;
	int iCubeJoinTicks;
	int iCubeFloatTicks;
	int iCubeFloatInterval;
	float fCubeFloatScale;
	int iTextOff;
	int iTextOff2;
	float fTextSpeed;

	// Model space: X to right, Y to top, Z goes into the screen, (0, 0, 0) at center of screen
	// Window space: X to right, Y to top, Z for depth testing only, (0, 0, z) at center of screen
	// Window: X to right, Y to bottom, no Z, (0, 0) at top left corner
	float fProjTheta; // Tilt from model +Y
	float fProjScale; // Window [0, iWndWidth] -> Window space [-fProjScale, fProjScale]
	point3_t ptCubeSize;
	std::array<point3_t, 4> ptCubePosInit;
	std::array<float, 4> fCubePosOff;
	std::array<float, 4> fCubeRotOff;

	int cubeTbdIdx[4];
	cube2_t cubeTbd[4];

	std::unique_ptr<Gdiplus::Pen> penEdge;
	std::unique_ptr<Gdiplus::Brush> brFill[4];

	std::unique_ptr<Gdiplus::Font> fntText;
	std::unique_ptr<Gdiplus::Brush> brText;

	std::unique_ptr<Gdiplus::Bitmap> bmBackground;
} g_ps{ 0 };

// ---

// Isometric projection from model space to window space
static point2d_t ProjectPoint(const point3_t& pt)
{
	float x, y, z;

	// Model space -> Window space
	// NOTE on minus sign: fProjTheta is of the opposite direction of normal angles
	x = pt[0]; // 1:1 on X direction
	y = sinf(-g_ps.fProjTheta) * pt[2] + cosf(-g_ps.fProjTheta) * pt[1];
	z = cosf(-g_ps.fProjTheta) * pt[2] - sinf(-g_ps.fProjTheta) * pt[1];

	// Window space -> Window
	x = x * g_ps.iWndWidth / (2 * g_ps.fProjScale) + g_ps.iWndWidth / 2.0f;
	y = -y * g_ps.iWndWidth / (2 * g_ps.fProjScale) + g_ps.iWndHeight / 2.0f;

	return { point2_t(x, y), z };
}

static face2d_t ProjectFace(const face3_t& face)
{
	face2_t out;
	float zsum = 0;
	for (int i = 0; i < 4; i++)
	{
		float z;
		std::tie(out[i], z) = ProjectPoint(face[i]);
		zsum += z;
	}
	return { out, zsum };
}

static cube2d_t ProjectCube(const cube3_t& cube)
{
	std::array<face2d_t, 6> faces;
	std::transform(cube.cbegin(), cube.cend(), faces.begin(), ProjectFace);
	std::sort(faces.begin(), faces.end(), [](auto& a, auto& b) { return std::get<1>(a) < std::get<1>(b); });

	cube2_t out;
	float zsum = 0;
	for (int i = 0; i < 6; i++)
	{
		float z;
		std::tie(out[i], z) = faces[i];
		zsum += z;
	}
	return { out, zsum };
}

// rot is on xOz plane
static cube3_t GenerateCube(const point3_t& origin, float rotoff, float rot)
{
	float w, h, l;
	w = g_ps.ptCubeSize[0];
	h = g_ps.ptCubeSize[1];
	l = g_ps.ptCubeSize[2];

	std::array<point3_t, 8> pts = { {
		{ 0, 0, 0 }, { w, 0, 0 }, { w, 0, l }, { 0, 0, l },
		{ 0, h, 0 }, { w, h, 0 }, { w, h, l }, { 0, h, l }
	} };
	for (auto& pt : pts)
	{
		float x, y, z;
		x = pt[0] + rotoff;
		y = pt[1];
		z = pt[2] + rotoff;

		pt[0] = origin[0] + x * cosf(rot) + z * sinf(rot);
		pt[1] = origin[1] + y;
		pt[2] = origin[2] - x * sinf(rot) + z * cosf(rot);
	}
	return { {
		{ pts[0], pts[1], pts[2], pts[3] },
		{ pts[0], pts[4], pts[5], pts[1] },
		{ pts[0], pts[4], pts[7], pts[3] },

		{ pts[6], pts[7], pts[4], pts[5] },
		{ pts[6], pts[2], pts[1], pts[5] },
		{ pts[6], pts[2], pts[3], pts[7] }
	} };
}

static void OnInit()
{
	RECT rc;
	GetClientRect(g_ps.hWnd, &rc);
	g_ps.iWndWidth = rc.right;
	g_ps.iWndHeight = rc.bottom;

	g_ps.iTicks = 0;
	g_ps.fBgSpeed = 1;
	g_ps.fCubeRotSpeed = -1.0 / 64;
	g_ps.iCubeJoinTicks = 1000;
	g_ps.iCubeFloatTicks = 120;
	g_ps.iCubeFloatInterval = 800;
	g_ps.fCubeFloatScale = 3;
	g_ps.iTextOff = 10;
	g_ps.iTextOff2 = 5;
	g_ps.fTextSpeed = 0.25;

	g_ps.fProjTheta = (float)M_PI / 8;
	g_ps.fProjScale = 2.5;
	g_ps.ptCubeSize = { 1.0, 0.5, 1.0 };
	g_ps.ptCubePosInit = { {
		{  6,  0,  0 }, // Purple
		{  0,  6,  0 }, // Blue
		{ -6,  0,  0 }, // Orange
		{  0, -5,  0 }, // Yellow
	} };
	g_ps.fCubePosOff = {
		0.0f,
		0.2f,
		0.0f,
		0.4f,
	};
	g_ps.fCubeRotOff = {
		0.15f,
		0,
		0.15f,
		0,
	};

	// ---

	g_ps.penEdge = std::make_unique<Gdiplus::Pen>(Gdiplus::Color(0, 0, 0), 4);
	g_ps.penEdge->SetMiterLimit(1.1f);

	int i = 0;
	for (auto rgb : { 0x02D740B1, 0x02BD7B42, 0x023173E7, 0x0200CEFF })
	{
		Gdiplus::Color clr;
		clr.SetFromCOLORREF(rgb);
		g_ps.brFill[i++] = std::make_unique<Gdiplus::SolidBrush>(clr);
	}

	//g_ps.fntText = std::make_unique<Gdiplus::Font>(L"System", 12);
	g_ps.fntText = std::make_unique<Gdiplus::Font>(L"System", 14, Gdiplus::FontStyleBold, Gdiplus::UnitPixel);
	g_ps.brText = std::make_unique<Gdiplus::SolidBrush>(Gdiplus::Color(0, 0, 0));

	// ---

	auto& bmiHeader = g_ps.pbmiBackground->bmiHeader;
	if (bmiHeader.biCompression == BI_RGB && bmiHeader.biBitCount <= 8 && bmiHeader.biClrUsed != 0)
	{
		Gdiplus::Bitmap bm(g_ps.pbmiBackground, &g_ps.pbmiBackground->bmiColors[bmiHeader.biClrUsed]);

		auto bmScaledWidth = (bm.GetWidth() - 40) * g_ps.iWndHeight / bm.GetHeight();
		g_ps.bmBackground = std::make_unique<Gdiplus::Bitmap>(bmScaledWidth, g_ps.iWndHeight);

		Gdiplus::Graphics gr(g_ps.bmBackground.get());
		gr.SetInterpolationMode(Gdiplus::InterpolationModeNearestNeighbor);
		gr.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
		gr.DrawImage(
			&bm,
			Gdiplus::RectF(0, 0, bmScaledWidth, g_ps.iWndHeight),
			Gdiplus::RectF(0, 0, bm.GetWidth() - 40, bm.GetHeight()),
			Gdiplus::UnitPixel
		);
	}
	else
	{
		g_ps.bmBackground = std::make_unique<Gdiplus::Bitmap>(g_ps.iWndWidth, g_ps.iWndHeight);
	}
}

static void OnFini()
{
	g_ps.penEdge.reset();
	for (auto& br : g_ps.brFill)
	{
		br.reset();
	}

	g_ps.fntText.reset();
	g_ps.brText.reset();

	g_ps.bmBackground.reset();
}

static void OnTick()
{
	g_ps.iTicks++;

	float fCubeRot = g_ps.iTicks * g_ps.fCubeRotSpeed;
	float fCubeJoinRatio = (float)min(g_ps.iTicks, g_ps.iCubeJoinTicks) / g_ps.iCubeJoinTicks; // 1 -> 0
	float fCubeFloatRatio = (float)max(g_ps.iTicks - (g_ps.iCubeJoinTicks + g_ps.iCubeFloatTicks), 0) / g_ps.iCubeFloatInterval; // 0 -> +inf
	std::array<point3_t, 4> ptCubePos = g_ps.ptCubePosInit;
	for (int i = 0; i < 4; i++)
	{
		for (auto& x : ptCubePos[i])
		{
			x *= 1 - fCubeJoinRatio;
		}
		ptCubePos[i][1] += powf(-1, roundf(fCubeFloatRatio)) * (fCubeFloatRatio - roundf(fCubeFloatRatio)) * g_ps.fCubeFloatScale;
		ptCubePos[i][1] += g_ps.fCubePosOff[i];
	}

	std::array<std::tuple<int, float>, 4> cubeIdx;
	for (int i = 0; i < 4; i++)
	{
		cube3_t cube3 = GenerateCube(ptCubePos[i], g_ps.fCubeRotOff[i], fCubeRot + (i + 1) * M_PI / 2);

		cube2_t cube;
		float z;
		std::tie(cube, z) = ProjectCube(cube3);

		g_ps.cubeTbd[i] = cube;
		cubeIdx[i] = { i, z };
	}
	std::sort(cubeIdx.begin(), cubeIdx.end(), [](auto& a, auto& b) { return std::get<1>(a) < std::get<1>(b); });
	std::transform(cubeIdx.begin(), cubeIdx.end(), g_ps.cubeTbdIdx, [](auto& x) { return std::get<0>(x); });
}

static void OnPaint(HDC hdc)
{
	Gdiplus::Graphics gr(hdc);
	gr.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias8x8);
	gr.SetTextRenderingHint(Gdiplus::TextRenderingHintSingleBitPerPixelGridFit);

	// --- Background

	Gdiplus::ImageAttributes imgAttr;
	imgAttr.SetWrapMode(Gdiplus::WrapModeTile);

	int iBgOff = g_ps.iTicks * g_ps.fBgSpeed;
	gr.DrawImage(
		g_ps.bmBackground.get(),
		Gdiplus::RectF(0, 0, g_ps.iWndWidth, g_ps.iWndHeight),
		Gdiplus::RectF(iBgOff, 0, g_ps.iWndWidth, g_ps.iWndHeight),
		Gdiplus::UnitPixel,
		&imgAttr
	);

	// --- Cubes

	for (auto idx : g_ps.cubeTbdIdx)
	{
		// Only draw the faces that could be shown
		for (int i = 3; i < 6; i++)
		{
			auto& face = g_ps.cubeTbd[idx][i];
			gr.FillPolygon(g_ps.brFill[idx].get(), face.data(), (int)face.size());
			gr.DrawPolygon(g_ps.penEdge.get(), face.data(), (int)face.size());
		}
	}

	// --- Text

	Gdiplus::StringFormat strFormat(Gdiplus::StringFormat::GenericDefault());
	strFormat.SetAlignment(Gdiplus::StringAlignmentCenter);
	strFormat.SetLineAlignment(Gdiplus::StringAlignmentNear);

	Gdiplus::RectF rc;
	gr.MeasureString(
		g_ps.wszText,
		-1,
		g_ps.fntText.get(),
		Gdiplus::RectF(0, 0, g_ps.iWndWidth, 1024),
		&strFormat,
		&rc,
		NULL,
		NULL
	);

	int iTextViewport = g_ps.iTextOff2 + g_ps.iWndHeight + g_ps.iTextOff;
	float fTextOff = iTextViewport - fmodf(g_ps.iTicks * g_ps.fTextSpeed, iTextViewport + rc.Height);
	gr.DrawString(
		g_ps.wszText,
		-1,
		g_ps.fntText.get(),
		Gdiplus::RectF(0, (int)fTextOff, g_ps.iWndWidth, 1024),
		&strFormat,
		g_ps.brText.get()
	);
}

// ---

static LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	switch (Msg)
	{
	case WM_CREATE:
	{
		g_ps.hWnd = hWnd;

#ifdef ENABLE_MIDI
		GetTempPathW(_countof(g_ps.wszMidiFileName), g_ps.wszMidiFileName);
		GetTempFileNameW(g_ps.wszMidiFileName, L"VBA", 0, g_ps.wszMidiFileName);

		HANDLE hFile;
		DWORD cbWritten;
		hFile = CreateFileW(g_ps.wszMidiFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY, NULL);
		WriteFile(hFile, g_ps.pvMidi, g_ps.cbMidi, &cbWritten, NULL);
		CloseHandle(hFile);

		MCI_OPEN_PARMSW openParams{ 0 };
		openParams.lpstrDeviceType = L"sequencer";
		openParams.lpstrElementName = g_ps.wszMidiFileName;
		if (!mciSendCommandW(0, MCI_OPEN, MCI_OPEN_TYPE | MCI_OPEN_ELEMENT, (DWORD_PTR)&openParams))
		{
			g_ps.wMidiDevice = openParams.wDeviceID;

			// Initial push to the play loop
			PostMessageW(hWnd, MM_MCINOTIFY, MCI_NOTIFY_SUCCESSFUL, g_ps.wMidiDevice);
		}
#endif

		Gdiplus::GdiplusStartupInput input;
		Gdiplus::GdiplusStartup(&g_ps.ulpGdiplusToken, &input, NULL);

		OnInit();
		SetTimer(hWnd, 12345, 25, NULL);

		return TRUE;
	}
	case WM_DESTROY:
	{
		KillTimer(hWnd, 12345);
		OnFini();

		if (g_ps.ulpGdiplusToken)
		{
			Gdiplus::GdiplusShutdown(g_ps.ulpGdiplusToken);
			g_ps.ulpGdiplusToken = 0;
		}
#ifdef ENABLE_MIDI
		if (g_ps.wMidiDevice)
		{
			mciSendCommandW(g_ps.wMidiDevice, MCI_CLOSE, 0, 0);
			g_ps.wMidiDevice = 0;
		}
		if (g_ps.wszMidiFileName[0])
		{
			DeleteFileW(g_ps.wszMidiFileName);
			g_ps.wszMidiFileName[0] = L'\0';
		}
#endif

		g_ps.hWnd = NULL;
		return TRUE;
	}
	case WM_TIMER:
	{
		if ((UINT_PTR)wParam == 12345)
		{
			OnTick();
			RedrawWindow(hWnd, NULL, NULL, RDW_INVALIDATE);
		}
		return TRUE;
	}
	case WM_PAINT:
	{
		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(hWnd, &ps);
		OnPaint(hdc);
		EndPaint(hWnd, &ps);
		return 0;
	}
#ifdef ENABLE_MIDI
	case MM_MCINOTIFY:
	{
		if (g_ps.wMidiDevice)
		{
			MCI_SEEK_PARMS seekParams{ 0 };
			mciSendCommandW(g_ps.wMidiDevice, MCI_SEEK, MCI_SEEK_TO_START, (DWORD_PTR)&seekParams);

			MCI_PLAY_PARMS playParams{ 0 };
			playParams.dwCallback = (DWORD_PTR)hWnd;
			mciSendCommandW(g_ps.wMidiDevice, MCI_PLAY, MCI_NOTIFY, (DWORD_PTR)&playParams);
		}
		return TRUE;
	}
#endif
	case WM_KEYDOWN:
	case WM_SYSKEYDOWN:
	{
		if (wParam == VK_ESCAPE)
		{
			PostMessageW(hWnd, WM_CLOSE, 0, 0);
		}
		break;
	}
	default:
		break;
	}
	return DefWindowProcW(hWnd, Msg, wParam, lParam);
}

// ---

static LPVOID GetResourcePtr(HINSTANCE hInstance, LPCWSTR lpName, LPCWSTR lpType, DWORD* pcbData)
{
	*pcbData = 0;

	HRSRC hrsrc = FindResourceW(hInstance, lpName, lpType);
	if (!hrsrc)
	{
		return NULL;
	}

	HGLOBAL hResData = LoadResource(hInstance, hrsrc);
	if (!hResData)
	{
		return NULL;
	}

	*pcbData = SizeofResource(hInstance, hrsrc);
	return LockResource(hResData);
}

__declspec(noinline)
static BOOL WINAPI FShowOfficeHelper()
{
	HINSTANCE hInstExe = GetModuleHandleW(NULL);

	if (!g_ps.fDrawnOnce)
	{
		WNDCLASSEXW cls{ sizeof(cls) };
		cls.lpszClassName = L"OfVbEg";
		cls.hInstance = hInstExe;
		cls.lpfnWndProc = WndProc;
		cls.style = CS_HREDRAW | CS_VREDRAW;
		cls.hIcon = NULL;
		cls.hIconSm = NULL;
		cls.hCursor = LoadCursorW(NULL, IDC_ARROW);
		cls.hbrBackground = NULL;
		g_ps.atomOfVbEg = RegisterClassExW(&cls);

		DWORD cbData;
		auto pData = GetResourcePtr(hInstExe, MAKEINTRESOURCEW(5434), RT_RCDATA, &cbData);
		memcpy(g_ps.wszText, pData, cbData);
		for (DWORD off = 0; off < cbData; off += sizeof(DWORD))
		{
			auto& x = *(PDWORD)((PBYTE)g_ps.wszText + off);
			x ^= 0x91779111;
			x &= 0x00FF00FF;
		}

		HINSTANCE hInstVbe7Intl = GetModuleHandleW(L"VBE7INTL.DLL");
		LoadStringW(hInstVbe7Intl, 13000 + 136, g_ps.wszTitle, _countof(g_ps.wszTitle));

		HINSTANCE hInstVbe7 = GetModuleHandleW(L"VBE7.DLL");
		g_ps.pvMidi = GetResourcePtr(hInstVbe7, MAKEINTRESOURCEW(5432), RT_RCDATA, &g_ps.cbMidi);
		g_ps.pbmiBackground = (LPBITMAPINFO)GetResourcePtr(hInstVbe7, MAKEINTRESOURCEW(5433), RT_BITMAP, &g_ps.cbBackground);

		g_ps.fDrawnOnce = TRUE;
	}

	if (!g_ps.hWnd)
	{
		CreateWindowExW(
			WS_EX_DLGMODALFRAME | WS_EX_COMPOSITED,
			MAKEINTATOM(g_ps.atomOfVbEg),
			g_ps.wszTitle,
			WS_VISIBLE | WS_CAPTION | WS_SYSMENU,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			300,
			400,
			g_hwndMainWindow,
			NULL,
			hInstExe,
			NULL
		);
	}

	return TRUE;
}

__declspec(noinline)
static BOOL WINAPI FIsOfficeCmd(LPCSTR szCaption)
{
    if (!szCaption || strlen(szCaption) != 15)
    {
        return FALSE;
    }

    auto pdwValues = (LPDWORD)szCaption;
    return (pdwValues[0] + pdwValues[1] + pdwValues[2] + pdwValues[3]) == 0xFC8AA51F
        && (pdwValues[0] & pdwValues[1] ^ pdwValues[2] & pdwValues[3]) == 0x20233041;
}

__declspec(noinline)
void vbofeg(const char* param)
{
    BYTE bKey = 0;
    bKey ^= (GetAsyncKeyState(VK_SHIFT) & 0x8000) >> 11;
    bKey ^= (GetAsyncKeyState(VK_OEM_3) & 0x8000) >> 15;

    CHAR szCaption[16];
    strncpy_s(szCaption, param, _TRUNCATE);
    for (char* p = szCaption; *p; p++)
    {
        *p ^= bKey;
    }

    if (FIsOfficeCmd(szCaption))
    {
        FShowOfficeHelper();
    }
    else
    {
        MessageBoxW(g_hwndMainWindow, L"... but can you catch the ti(l)de?", L"The twilight awaits...", MB_ICONINFORMATION);
    }
}

