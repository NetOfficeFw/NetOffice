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

HRESULT DllRegister(HINSTANCE module, LPCWSTR officeApplication, DWORD addinLoadBehavior, LPCWSTR progId, LPCWSTR classId, LPCWSTR description, RegisterMode mode);

HRESULT DllUnregister(LPCWSTR officeApplication, LPCWSTR progId, LPCWSTR classId, RegisterMode mode);

//HRESULT RegisterCOMAddin(LPCWSTR pszOfficeApp, LPCWSTR pszProgID, LPCWSTR pszFriendlyName, DWORD dwStartupContext);
//
//HRESULT UnRegisterCOMAddin(LPCWSTR pszOfficeApp, LPCWSTR pszProgID);
//
//BOOL SetKeyAndValue(HKEY hKeyRoot, LPCWSTR pszPath, LPCWSTR pszSubkey, LPCWSTR pszValue);
//
//LONG RecursiveDeleteKey(HKEY hKeyParent, LPCWSTR pszKeyChild);
