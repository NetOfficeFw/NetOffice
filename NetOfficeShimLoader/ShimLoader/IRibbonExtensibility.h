#pragma once
#include "stdafx.h"

__interface __declspec(uuid("000c0396-0000-0000-c000-000000000046"))
	IRibbonExtensibility : IDispatch
{
	STDMETHOD(GetCustomUI)(THIS_ BSTR RibbonID, BSTR* RibbonXml) PURE;
};
// 000c0396-0000-0000-c000-000000000046
static const GUID IID_IRibbonExtensibility = __uuidof(IRibbonExtensibility);