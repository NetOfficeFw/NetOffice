#pragma once
#include "stdAfx.h"

//
// Definition of Aggregators
//

__interface __declspec(uuid("E8E14A9B-6FB4-45A6-BFF2-47610F68D075"))
	IOuterComAggregator : public IUnknown
{
	//
	// Inner managed aggreator call this method to publish the managed addin instance to the shim
	//
	// param pUnkInner: managed addin instance
	HRESULT __stdcall SetInnerAddin(IUnknown* pUnkInner);
};

namespace NetOffice_Tools_Isolation
{
	//
	// Represents the inner managed aggregator in NetOffice Core Assembly
	//
	__interface __declspec(uuid("FBA7450D-B6E0-4E5C-908D-396BEFFC1D9B"))
		IManagedInnerAggregator : public IUnknown
	{
		//
		// Creates a managed instance by name and call SetInnerAddin for pOuterObject argument
		//
		// param assemblyName: name or strong name from the assembly where the managed addin type is located
		// param fullQualifiedTypeName: full qualified name of the managed addin type
		// param pOuterObject: outer aggregator in order to call SetInnerAddin to publish the newly created instance
		HRESULT __stdcall CreateAggregatedInstance(BSTR assemblyName, BSTR fullQualifiedTypeName, IOuterComAggregator* pOuterObject);
	};
}
