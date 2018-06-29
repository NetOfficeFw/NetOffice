#pragma once
#include "stdAfx.h"

//
// Definition of Aggregators
//

__interface __declspec(uuid("E8E14A9B-6FB4-45A6-BFF2-47610F68D075"))
	IOuterComAggregator : public IUnknown
{
	HRESULT __stdcall SetInnerAddin(IUnknown *pUnkInner);
};

namespace NetOffice_Tools_Isolation
{
	__interface __declspec(uuid("FBA7450D-B6E0-4E5C-908D-396BEFFC1D9B"))
		IManagedInnerAggregator : public IUnknown
	{
		HRESULT __stdcall CreateAggregatedInstance(BSTR assemblyName, BSTR bstrTypeName, IOuterComAggregator* pOuterObject);
	};
}
