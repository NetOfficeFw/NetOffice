#include <windows.h>
#include <oleauto.h>
#include <atomic>
#include <new>
#include "LifetimeFixture_h.h"

namespace
{
    std::atomic<long> g_liveObjects{ 0 };
    std::atomic<long> g_constructCount{ 0 };
    std::atomic<long> g_destroyCount{ 0 };
    std::atomic<long> g_addRefCount{ 0 };
    std::atomic<long> g_releaseCount{ 0 };
    std::atomic<long> g_invokeCount{ 0 };
    std::atomic<long> g_serverLocks{ 0 };
    std::atomic<long> g_liveClassFactories{ 0 };

    class NativeProbeObject final : public IProbeObject
    {
    public:
        NativeProbeObject() : _references(1)
        {
            ++g_liveObjects;
            ++g_constructCount;
        }

        ~NativeProbeObject()
        {
            --g_liveObjects;
            ++g_destroyCount;
        }

        HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** value) override
        {
            if (value == nullptr)
                return E_POINTER;

            *value = nullptr;
            if (iid == IID_IUnknown || iid == IID_IDispatch || iid == __uuidof(IProbeObject))
                *value = static_cast<IProbeObject*>(this);

            if (*value == nullptr)
                return E_NOINTERFACE;

            AddRef();
            return S_OK;
        }

        ULONG STDMETHODCALLTYPE AddRef() override
        {
            ++g_addRefCount;
            return static_cast<ULONG>(InterlockedIncrement(&_references));
        }

        ULONG STDMETHODCALLTYPE Release() override
        {
            ++g_releaseCount;
            const ULONG remaining = static_cast<ULONG>(InterlockedDecrement(&_references));
            if (remaining == 0)
                delete this;
            return remaining;
        }

        HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* count) override
        {
            if (count == nullptr)
                return E_POINTER;
            *count = 0;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override
        {
            return E_NOTIMPL;
        }

        HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID iid, LPOLESTR* names, UINT count, LCID, DISPID* ids) override
        {
            if (iid != IID_NULL)
                return DISP_E_UNKNOWNINTERFACE;
            if (names == nullptr || ids == nullptr)
                return E_POINTER;

            for (UINT index = 0; index < count; ++index)
            {
                if (_wcsicmp(names[index], L"Identity") == 0)
                    ids[index] = 1;
                else if (_wcsicmp(names[index], L"Ping") == 0)
                    ids[index] = 2;
                else if (_wcsicmp(names[index], L"Self") == 0)
                    ids[index] = 3;
                else
                    return DISP_E_UNKNOWNNAME;
            }
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE Invoke(DISPID member, REFIID iid, LCID, WORD flags, DISPPARAMS*, VARIANT* result, EXCEPINFO*, UINT*) override
        {
            if (iid != IID_NULL)
                return DISP_E_UNKNOWNINTERFACE;
            if (result == nullptr)
                return E_POINTER;

            VariantInit(result);
            ++g_invokeCount;

            switch (member)
            {
            case 1:
                if ((flags & DISPATCH_PROPERTYGET) == 0)
                    return DISP_E_MEMBERNOTFOUND;
                result->vt = VT_I4;
                result->lVal = 4242;
                return S_OK;
            case 2:
                if ((flags & DISPATCH_METHOD) == 0)
                    return DISP_E_MEMBERNOTFOUND;
                result->vt = VT_BSTR;
                result->bstrVal = SysAllocString(L"pong");
                return result->bstrVal == nullptr ? E_OUTOFMEMORY : S_OK;
            case 3:
                if ((flags & DISPATCH_PROPERTYGET) == 0)
                    return DISP_E_MEMBERNOTFOUND;
                result->vt = VT_DISPATCH;
                result->pdispVal = static_cast<IDispatch*>(this);
                AddRef();
                return S_OK;
            default:
                return DISP_E_MEMBERNOTFOUND;
            }
        }

        HRESULT STDMETHODCALLTYPE get_Identity(LONG* value) override
        {
            if (value == nullptr)
                return E_POINTER;
            *value = 4242;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE Ping(BSTR* value) override
        {
            if (value == nullptr)
                return E_POINTER;
            *value = SysAllocString(L"pong");
            return *value == nullptr ? E_OUTOFMEMORY : S_OK;
        }

        HRESULT STDMETHODCALLTYPE get_Self(IProbeObject** value) override
        {
            if (value == nullptr)
                return E_POINTER;
            *value = this;
            AddRef();
            return S_OK;
        }

    private:
        volatile long _references;
    };

    class ProbeClassFactory final : public IClassFactory
    {
    public:
        ProbeClassFactory() : _references(1) { ++g_liveClassFactories; }
        ~ProbeClassFactory() { --g_liveClassFactories; }

        HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** value) override
        {
            if (value == nullptr)
                return E_POINTER;
            *value = nullptr;
            if (iid == IID_IUnknown || iid == IID_IClassFactory)
                *value = static_cast<IClassFactory*>(this);
            if (*value == nullptr)
                return E_NOINTERFACE;
            AddRef();
            return S_OK;
        }

        ULONG STDMETHODCALLTYPE AddRef() override
        {
            return static_cast<ULONG>(InterlockedIncrement(&_references));
        }

        ULONG STDMETHODCALLTYPE Release() override
        {
            const ULONG remaining = static_cast<ULONG>(InterlockedDecrement(&_references));
            if (remaining == 0)
                delete this;
            return remaining;
        }

        HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid, void** value) override
        {
            if (outer != nullptr)
                return CLASS_E_NOAGGREGATION;
            if (value == nullptr)
                return E_POINTER;

            *value = nullptr;
            NativeProbeObject* instance = new (std::nothrow) NativeProbeObject();
            if (instance == nullptr)
                return E_OUTOFMEMORY;

            const HRESULT result = instance->QueryInterface(iid, value);
            instance->Release();
            return result;
        }

        HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override
        {
            if (lock)
                ++g_serverLocks;
            else
                --g_serverLocks;
            return S_OK;
        }

    private:
        volatile long _references;
    };
}

extern "C" __declspec(dllexport) void __stdcall LifetimeFixture_ResetTelemetry()
{
    g_constructCount = 0;
    g_destroyCount = 0;
    g_addRefCount = 0;
    g_releaseCount = 0;
    g_invokeCount = 0;
}

extern "C" __declspec(dllexport) long __stdcall LifetimeFixture_GetLiveObjectCount() { return g_liveObjects.load(); }
extern "C" __declspec(dllexport) long __stdcall LifetimeFixture_GetConstructCount() { return g_constructCount.load(); }
extern "C" __declspec(dllexport) long __stdcall LifetimeFixture_GetDestroyCount() { return g_destroyCount.load(); }
extern "C" __declspec(dllexport) long __stdcall LifetimeFixture_GetAddRefCount() { return g_addRefCount.load(); }
extern "C" __declspec(dllexport) long __stdcall LifetimeFixture_GetReleaseCount() { return g_releaseCount.load(); }
extern "C" __declspec(dllexport) long __stdcall LifetimeFixture_GetInvokeCount() { return g_invokeCount.load(); }

STDAPI DllGetClassObject(REFCLSID clsid, REFIID iid, void** value)
{
    if (clsid != __uuidof(ProbeObject))
        return CLASS_E_CLASSNOTAVAILABLE;

    ProbeClassFactory* factory = new (std::nothrow) ProbeClassFactory();
    if (factory == nullptr)
        return E_OUTOFMEMORY;

    const HRESULT result = factory->QueryInterface(iid, value);
    factory->Release();
    return result;
}

STDAPI DllCanUnloadNow()
{
    return g_liveObjects.load() == 0 && g_liveClassFactories.load() == 0 && g_serverLocks.load() == 0 ? S_OK : S_FALSE;
}

BOOL APIENTRY DllMain(HMODULE, DWORD, LPVOID)
{
    return TRUE;
}
