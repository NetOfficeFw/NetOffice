#pragma once
#include "stdafx.h"
#include "DllRegisterMode.hpp"

namespace ShimLoader_Register32
{
	//
	// Register Shim as 32 Bit Build on a 32 Bit System
	//
	HRESULT DllRegister(HINSTANCE module, LPCWSTR officeApplication, DWORD addinLoadBehavior, DWORD addinCommandLineSafe, LPCWSTR progId, LPCWSTR classId, LPCWSTR friendlyName, LPCWSTR description, LPCWSTR version, RegisterMode mode);

	//
	// Unregister Shim as 32 Bit Build on a 32 Bit System
	//
	HRESULT DllUnregister(LPCWSTR officeApplication, LPCWSTR progId, LPCWSTR classId, LPCWSTR version, RegisterMode mode);
}
