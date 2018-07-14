#pragma once
#include "stdafx.h"
#include "ShimProxy.h"

extern HANDLE		_thread;
extern HINSTANCE	_module;
extern void IncComponents(LPCWSTR type);
extern void DecComponents(LPCWSTR type);

namespace NetOffice_ShimLoader
{
	// TODO: rename this class because its construct
	// all COM instances per design, not just the ShimProxy
	// in fact yes its construct only the ShimProxy but not from a general design point of view
	class ShimProxyFactory : public IClassFactory
	{

	public:

		// Ctor, Dtor
		ShimProxyFactory();
		virtual ~ShimProxyFactory();

		// IClassFactory Implementation
		STDMETHODIMP  CreateInstance(LPUNKNOWN punk, REFIID riid, void** ppv);
		STDMETHODIMP  LockServer(BOOL fLock);

		// IUnknown Implementation
		STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

	private:

		ULONG           _refCount;
	};
}
