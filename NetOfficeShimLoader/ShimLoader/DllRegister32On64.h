#pragma once
#include "stdafx.h"
#include "DllRegisterMode.hpp"

using namespace NetOffice_ShimLoader;

namespace NetOffice_ShimLoader_Register32On64
{
	//
	// Register Shim as 32 Bit Build on a 64 Bit System
	//
	HRESULT DllRegister(HINSTANCE module, LPCWSTR officeApplications[], DWORD addinLoadBehavior, DWORD addinCommandLineSafe, WCHAR* progId, WCHAR* classId, WCHAR* friendlyName, WCHAR* description, WCHAR* version, RegisterMode mode, BOOL addinRegistration);

	//
	// Unregister Shim as 32 Bit Build on a 64 Bit System
	//
	HRESULT DllUnregister(LPCWSTR officeApplications[], WCHAR* progId, WCHAR* classId, WCHAR* version, RegisterMode mode, BOOL addinRegistration);
}
