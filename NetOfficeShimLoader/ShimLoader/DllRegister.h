#pragma once
#include "stdafx.h"

/*
Determines the Assembly Registration Mode
*/
enum RegisterMode
{
	// Component and Addin is registered per Machine
	System = 0,

	// Component and Addin is registered per current User
	User = 1,

	// Component is registered per Machine
	// Addin is registered per current User
	SystemComponentAndUserAddin = 2
};

HRESULT DllRegister(HINSTANCE module, LPCWSTR officeApplication, DWORD addinLoadBehavior, DWORD addinCommandLineSafe, LPCWSTR progId, LPCWSTR classId, LPCWSTR friendlyName, LPCWSTR description, LPCWSTR version, RegisterMode mode);

HRESULT DllUnregister(LPCWSTR officeApplication, LPCWSTR progId, LPCWSTR classId, LPCWSTR version, RegisterMode mode);
