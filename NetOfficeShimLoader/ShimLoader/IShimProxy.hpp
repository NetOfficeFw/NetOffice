#pragma once
#include "stdAfx.h"

namespace NetOffice_ShimLoader
{
	//
	// Represents internal tasks of an outer aggregator aka Shim that handle update/reload possibilites
	// This representation is not provided to the inner addin
	//
	__interface __declspec(uuid("D3614A78-BA1D-49B7-BC02-991762879CA3")) IShimProxy
	{
		//
		// Determines the CLR is currently loaded
		//
		BOOL STDMETHODCALLTYPE IsCLRLoaded();

		//
		// Loads the CLR and its managed inner addin
		//
		// Fail: CLR is already loaded
		//
		STDMETHODIMP LoadCLR();

		//
		// Unloads the CLR and its managed inner addin
		//
		// Fail: CLR is already unloaded
		//
		//
		STDMETHODIMP UnloadCLR();

		//
		// Determines a reload thread is currently in progress
		//
		BOOL IsReloadThreadInProgress();

		//
		// Determines an async reload is currently in progress
		//
		BOOL IsAsyncReloadThreadInProgress();

		//
		// Unloads the CLR if necessary and loads the CLR and its managed addin (again)
		//
		// Param "async": Execute action in a worker thread. This must be used if the incoming call comes from the managed addin because
		// we are inside the (CLR)tread this method is going to kill and the CLR Hosting API doesnt like that(for a good reason).
		//
		// Fail: Another reload is currently in progress
		//
		STDMETHODIMP ReloadCLR(BOOL async);

		//
		// Free current ReloadCLR Thread
		// Failed if no reload thread is currently there
		//
		// Fail: No reload thread available
		//
		STDMETHODIMP CloseReloadThread();

		//
		// Reassign managed inner pointers for example IRibbonExtensibility or ICustomTaskPaneConsumer
		// to make sure application is calling valid pointers after reload CLR.
		//
		//
		STDMETHODIMP AssignInnerPointers();

		//
		//
		//
		STDMETHODIMP LoadUpdateHandler();

		//
		//
		//
		STDMETHODIMP Update(BOOL async);
	};
	// D3614A78-BA1D-49B7-BC02-991762879CA3
	static const GUID IID_IShimProxy = __uuidof(IShimProxy);
}
