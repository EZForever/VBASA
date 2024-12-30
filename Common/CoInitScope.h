#pragma once

#include <objbase.h>

#include <comdef.h>

class CoInitScope final
{
private:
    CoInitScope(CoInitScope&&) = delete;

public:
    inline CoInitScope(COINIT dwCoInit = COINIT_MULTITHREADED)
    {
        auto hr = CoInitializeEx(NULL, dwCoInit);
        if (FAILED(hr))
            _com_raise_error(hr);
    }

    inline ~CoInitScope()
    {
        CoUninitialize();
    }
};

