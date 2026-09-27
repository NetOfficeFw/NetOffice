using System.Collections.Immutable;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Runtime.Versioning;
namespace NetOffice.CodeGen.TypeLib;

/// <summary>REGKIND values accepted by LoadTypeLibEx.</summary>
public enum RegistryImportKind
{
    Default = 0,
    Register = 1,
    None = 2
}

/// <summary>Minimal native boundary, injectable so fixture imports never need Office or COM.</summary>
public interface INativeTypeLibApi
{
    int LoadTypeLibEx(string path, RegistryImportKind kind, out ITypeLib? typeLib);
    void Release(object? comObject);
}
/// <summary>Windows implementation of the native loader.</summary>
[SupportedOSPlatform("windows")]
public sealed class WindowsNativeTypeLibApi : INativeTypeLibApi
{
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern int NativeLoadTypeLibEx(string szFile, RegistryImportKind regkind, out ITypeLib? pptLib);
    public int LoadTypeLibEx(string path, RegistryImportKind kind, out ITypeLib? typeLib)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        return NativeLoadTypeLibEx(path, kind, out typeLib);
    }

    public void Release(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
            Marshal.FinalReleaseComObject(comObject);
    }
}

/// <summary>Decoder boundary separating COM descriptors from immutable observations.</summary>
public interface ITypeLibDescriptorDecoder
{
    TypeLibObservation Decode(ITypeLib typeLib, TypeLibProvenance provenance);
}

/// <summary>COM descriptor lifetime boundary used by the default decoder.</summary>
public interface IComDescriptorReleaser
{
    void ReleaseLibraryAttributes(ITypeLib typeLib, IntPtr attributes);
    void ReleaseTypeAttributes(ITypeInfo typeInfo, IntPtr attributes);
    void ReleaseFunctionDescriptor(ITypeInfo typeInfo, IntPtr descriptor);
    void ReleaseVariableDescriptor(ITypeInfo typeInfo, IntPtr descriptor);
    void ReleaseComObject(object? value);
}

/// <summary>Default descriptor releaser; every native descriptor is released by its owning COM object.</summary>
[SupportedOSPlatform("windows")]
public sealed class ComDescriptorReleaser : IComDescriptorReleaser
{
    public void ReleaseLibraryAttributes(ITypeLib typeLib, IntPtr attributes)
    {
        if (attributes != IntPtr.Zero) typeLib.ReleaseTLibAttr(attributes);
    }

    public void ReleaseTypeAttributes(ITypeInfo typeInfo, IntPtr attributes)
    {
        if (attributes != IntPtr.Zero) typeInfo.ReleaseTypeAttr(attributes);
    }

    public void ReleaseFunctionDescriptor(ITypeInfo typeInfo, IntPtr descriptor)
    {
        if (descriptor != IntPtr.Zero) typeInfo.ReleaseFuncDesc(descriptor);
    }

    public void ReleaseVariableDescriptor(ITypeInfo typeInfo, IntPtr descriptor)
    {
        if (descriptor != IntPtr.Zero) typeInfo.ReleaseVarDesc(descriptor);
    }

    public void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}

/// <summary>Load and decode one typelib with REGKIND_NONE and deterministic file provenance.</summary>
public sealed class TypeLibLoader
{
    private readonly INativeTypeLibApi _native;
    private readonly ITypeLibDescriptorDecoder _decoder;

    public TypeLibLoader(INativeTypeLibApi native, ITypeLibDescriptorDecoder decoder)
    {
        _native = native ?? throw new ArgumentNullException(nameof(native));
        _decoder = decoder ?? throw new ArgumentNullException(nameof(decoder));
    }

    public TypeLibObservation Load(string path, CurrentChannelMetadata? currentChannel = null, string? importerVersion = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        var bytes = File.ReadAllBytes(path);
        var digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        var provenance = new TypeLibProvenance(TypeLibSourceKind.File, Path.GetFullPath(path), digest, importerVersion, currentChannel, null);
        ITypeLib? typeLib = null;
        var hresult = _native.LoadTypeLibEx(path, RegistryImportKind.None, out typeLib);
        if (hresult < 0 || typeLib is null)
            throw new TypeLibLoadException(path, hresult);

        try
        {
            return _decoder.Decode(typeLib, provenance);
        }
        finally
        {
            _native.Release(typeLib);
        }
    }
}

/// <summary>Failure returned by LoadTypeLibEx, including the native HRESULT.</summary>
public sealed class TypeLibLoadException : IOException
{
    public TypeLibLoadException(string path, int hresult)
        : base($"LoadTypeLibEx(REGKIND_NONE) failed for '{path}' with HRESULT 0x{hresult:X8}.")
    {
        Path = path;
        HResult = hresult;
    }

    public string Path { get; }
}
/// <summary>Descriptor decoding failed while resolving a native reference.</summary>
public sealed class TypeLibDescriptorException : Exception
{
    public TypeLibDescriptorException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>Fixture-friendly decoder that accepts already decoded immutable observations.</summary>
public sealed class ObservationDecoder : ITypeLibDescriptorDecoder
{
    private readonly Func<ITypeLib, TypeLibProvenance, TypeLibObservation> _decode;

    public ObservationDecoder(Func<ITypeLib, TypeLibProvenance, TypeLibObservation> decode)
    {
        _decode = decode ?? throw new ArgumentNullException(nameof(decode));
    }

    public TypeLibObservation Decode(ITypeLib typeLib, TypeLibProvenance provenance)
    {
        ArgumentNullException.ThrowIfNull(typeLib);
        ArgumentNullException.ThrowIfNull(provenance);
        return _decode(typeLib, provenance);
    }
}
