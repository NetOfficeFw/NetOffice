#pragma once
#include "StdAfx.h"
#include "metahost.h"
#include "atlbase.h"
#include "strsafe.h"

extern HINSTANCE _module;
extern ULONG _components;
extern ULONG _locks;

class ShimArguments
{
public:
	ShimArguments();
	~ShimArguments();

	HRESULT Load();
	HRESULT Unload();

protected:

	HRESULT AppendPath(LPWSTR pszPath, LPCWSTR pszMore);

};
