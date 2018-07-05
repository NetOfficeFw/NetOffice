#pragma once
#include "stdafx.h"
#include "DllRegisterMode.hpp"

namespace ShimLoader_Register32On64
{
	HRESULT DllRegister(HINSTANCE module, LPCWSTR officeApplication, DWORD addinLoadBehavior, DWORD addinCommandLineSafe, LPCWSTR progId, LPCWSTR classId, LPCWSTR friendlyName, LPCWSTR description, LPCWSTR version, RegisterMode mode);

	HRESULT DllUnregister(LPCWSTR officeApplication, LPCWSTR progId, LPCWSTR classId, LPCWSTR version, RegisterMode mode);
}
