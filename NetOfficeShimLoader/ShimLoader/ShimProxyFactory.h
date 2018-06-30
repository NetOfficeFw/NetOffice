#pragma once

class ShimProxyFactory : public IClassFactory
{

public:

	// Ctor, Dtor
	ShimProxyFactory();
	~ShimProxyFactory();

	// IUnknown Implementation
	STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	// IClassFactory Implementation
	STDMETHODIMP  CreateInstance(LPUNKNOWN punk, REFIID riid, void** ppv);
	STDMETHODIMP  LockServer(BOOL fLock);

private:

	ULONG           _refCount;
};

