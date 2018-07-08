#pragma once
#include "stdAfx.h"

namespace NetOffice_Tools_Isolation
{
	//
	//  Represents an outer shim host for a managed update handler
	//
	__interface __declspec(uuid("EF2F0985-2D4F-45AA-ADB6-510271A6EFC3"))
		IShimUpdateHost : public IUnknown
	{

		STDMETHODIMP SetCustomData(BSTR custom);

		//
		// Signalize the update is done
		// The host want unload the update handler and reload the managed addin
		//
		STDMETHODIMP Done();
	};
	// EF2F0985-2D4F-45AA-ADB6-510271A6EFC3
	static const GUID IID_IShimUpdateHost = __uuidof(IShimUpdateHost);

	//
	// Represents an outer shim host for a managed update handler
	//
	__interface __declspec(uuid("E20B53FD-03C8-4977-8725-7E0C89657960"))
		IOuterUpdateAggregator : public IUnknown
	{
		//
		// Inner update aggreator call this method to publish the managed addin instance to the shim
		//
		// param "pUnkInner": managed addin instance
		STDMETHODIMP SetInnerHandler(IUnknown* innerHandler);
	};
	// E20B53FD-03C8-4977-8725-7E0C89657960
	static const GUID IID_IOuterUpdateAggregator = __uuidof(IOuterUpdateAggregator);

	//
	// Represents an inner aggregator by a managed update handler
	//
	__interface __declspec(uuid("F249CE55-8BC5-44AB-81A0-1040998B6CD5"))
		IManagedInnerUpdateAggregator : public IUnknown
	{
		//
		// Creates a managed instance by name and call SetInnerHandler for pOuterObject argument
		//
		// param "assemblyName": name or strong name from the assembly where the managed addin type is located
		// param "fullQualifiedTypeName": full qualified name of the managed addin type
		// param "pOuterObject": outer aggregator in order to call SetInnerAddin to publish the newly created instance
		// param "shimHost" : outer shim host
		STDMETHODIMP CreateAggregatedInstance(BSTR assemblyName, BSTR fullQualifiedTypeName, IOuterUpdateAggregator* pOuterObject, IShimUpdateHost* shimHost);
	};
	// F249CE55-8BC5-44AB-81A0-1040998B6CD5
	static const GUID IID_IManagedInnerUpdateAggregator = __uuidof(IManagedInnerUpdateAggregator);

	//
	// To implement by a managed instance to recieve an outer shim host and execute update
	//
	__interface __declspec(uuid("BA23F519-0F53-4EC7-A416-2681BE22150F"))
		IManagedInnerUpdateHandler : public IUnknown
	{
		//
		// Set an outer parent shim host to the update handler
		//
		// param "shim": outer parent shim
		//
		STDMETHODIMP SetParent(IShimUpdateHost* shim);

		//
		// Set custom data from addin instance to the update handler
		//
		// param "custom": custom data as any coming from the addin that has requested the update
		//
		STDMETHODIMP SetCustomData(BSTR custom);

		//
		//
		// Set the host application to the update handler if OnConnection is already passed
		//
		//
		// param "application": host application
		//
		STDMETHODIMP SetApplication(IDispatch* application);

		//
		// Determines the update handler supports direct execution
		// Outer update aggregator want execute and close the handler if execution is supported
		//
		// param "canExecute": true if direct execution is supported, otherwise false
		//
		STDMETHODIMP CanExecute(BOOL* canExecute);

		//
		// Execute the handler
		//
		STDMETHODIMP Execute();

		//
		// Called before the outer shim is unload the AppDomain
		//
		STDMETHODIMP Close();
	};
	// BA23F519-0F53-4EC7-A416-2681BE22150F
	static const GUID IID_IManagedInnerUpdateHandler = __uuidof(IManagedInnerUpdateHandler);

	//
	// Represents an umanaged shim that is available for a managed inner addin
	//
	__interface __declspec(uuid("F7BCF161-FCB2-4880-9C33-78C456B1F291"))
		IShimHost : public IUnknown
	{
		//
		// Determines the shim is available
		// The shim is never available if blind aggregation is enabled
		//
		STDMETHODIMP IsAvailable(BOOL* available);

		//
		// Recreate the managed appdomain and create a new instance of the managed addin
		//
		STDMETHODIMP Reload();

		//
		//
		//
		STDMETHODIMP Update(BSTR custom);
	};
	// F7BCF161-FCB2-4880-9C33-78C456B1F291
	static const GUID IID_IShimHost = __uuidof(IShimHost);

	//
	// Represents a managed inner addin that is a child from a shim
	//
	__interface __declspec(uuid("EF261BCD-3078-459E-9448-13845BEED136"))
		IManagedInnerAddin : public IUnknown
	{
		//
		// Set an outer parent shim host to the addin
		//
		// param "shim": outer parent shim
		//
		STDMETHODIMP SetParent(IShimHost* shim);

		//
		// Notifies the adddin that its loaded by an update request
		//
		// param "custom": custom data given in prev managed addin instance (possibly modified by an update handler)
		//
		//
		STDMETHODIMP ReloadNotification(BSTR custom);
	};
	// F7BCF161-FCB2-4880-9C33-78C456B1F291
	static const GUID IID_IManagedInnerAddin = __uuidof(IManagedInnerAddin);

	//
	// Definition of Aggregators
	//
	__interface __declspec(uuid("E8E14A9B-6FB4-45A6-BFF2-47610F68D075"))
		IOuterComAggregator : public IUnknown
	{
		//
		// Inner managed aggreator call this method to publish the managed addin instance to the shim
		//
		// param "pUnkInner": managed addin instance
		STDMETHODIMP SetInnerAddin(IUnknown* pUnkInner);
	};
	// E8E14A9B-6FB4-45A6-BFF2-47610F68D075
	static const GUID IID_IOuterComAggregator = __uuidof(IOuterComAggregator);

	//
	// Represents the inner managed aggregator in NetOffice Core Assembly
	//
	__interface __declspec(uuid("FBA7450D-B6E0-4E5C-908D-396BEFFC1D9B"))
		IManagedInnerComAggregator : public IUnknown
	{
		//
		// Creates a managed instance by name and call SetInnerAddin for pOuterObject argument
		//
		// param "assemblyName": name or strong name from the assembly where the managed addin type is located
		// param "fullQualifiedTypeName": full qualified name of the managed addin type
		// param "pOuterObject": outer aggregator in order to call SetInnerAddin to publish the newly created instance
		// param "pOuterUpdateObject" : outer shim host
		STDMETHODIMP CreateAggregatedInstance(BSTR assemblyName, BSTR fullQualifiedTypeName, IOuterComAggregator* pOuterObject, IShimHost* pOuterUpdateObject);
	};
	// FBA7450D-B6E0-4E5C-908D-396BEFFC1D9B
	static const GUID IID_IManagedInnerComAggregator = __uuidof(IManagedInnerComAggregator);
}
