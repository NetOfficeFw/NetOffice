#pragma once
#include "stdAfx.h"

//
// Represents an outer aggregator by a addin that handle update/reload possibilites
//
__interface __declspec(uuid("D3614A78-BA1D-49B7-BC02-991762879CA3"))
	IShimProxy : public IUnknown
{
	//
	//
	//
	BOOL STDMETHODCALLTYPE IsCLRLoaded();

	//
	//
	//
	STDMETHODIMP ReloadCLR();

	//
	//
	//
	STDMETHODIMP UnloadCLR();
};
static const GUID IID_IShimProxy = __uuidof(IShimProxy);