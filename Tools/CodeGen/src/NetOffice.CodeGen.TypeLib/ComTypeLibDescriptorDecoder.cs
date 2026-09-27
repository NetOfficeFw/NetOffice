using System.Collections.Immutable;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
namespace NetOffice.CodeGen.TypeLib;

/// <summary>Decodes the standard COM typelib descriptors without retaining native pointers.</summary>
[SupportedOSPlatform("windows")]
public sealed class ComTypeLibDescriptorDecoder : ITypeLibDescriptorDecoder
{
    private readonly IComDescriptorReleaser _releaser;

    public ComTypeLibDescriptorDecoder(IComDescriptorReleaser? releaser = null)
    {
        _releaser = releaser ?? new ComDescriptorReleaser();
    }

    public TypeLibObservation Decode(ITypeLib typeLib, TypeLibProvenance provenance)
    {
        ArgumentNullException.ThrowIfNull(typeLib);
        ArgumentNullException.ThrowIfNull(provenance);

        IntPtr libraryAttributes = IntPtr.Zero;
        try
        {
            typeLib.GetLibAttr(out libraryAttributes);
            var attributes = Marshal.PtrToStructure<TYPELIBATTR>(libraryAttributes);
            var identity = new TypeLibIdentity(attributes.guid, unchecked((ushort)attributes.wMajorVerNum), unchecked((ushort)attributes.wMinorVerNum));
            GetDocumentation(typeLib, -1, out var name, out _);
            var types = ImmutableArray.CreateBuilder<TypeObservation>(typeLib.GetTypeInfoCount());
            for (var index = 0; index < typeLib.GetTypeInfoCount(); index++)
            {
                typeLib.GetTypeInfo(index, out var typeInfo);
                if (typeInfo is null) continue;
                try
                {
                    types.Add(ReadType(typeInfo, index));
                }
                finally
                {
                    _releaser.ReleaseComObject(typeInfo);
                }
            }

            var references = types
                .SelectMany(x => x.References)
                .Select(x => x.Library)
                .Distinct()
                .OrderBy(x => x.LibraryId)
                .ThenBy(x => x.MajorVersion)
                .ThenBy(x => x.MinorVersion)
                .ThenBy(x => x.Name, StringComparer.Ordinal)
                .ToImmutableArray();
            return new TypeLibObservation(
                TypeLibObservationSchema.Version,
                identity,
                name,
                attributes.lcid,
                (int)attributes.syskind,
                provenance,
                types.ToImmutable(),
                references,
                ImmutableDictionary<string, string>.Empty);
        }
        finally
        {
            _releaser.ReleaseLibraryAttributes(typeLib, libraryAttributes);
        }
    }

    private TypeObservation ReadType(ITypeInfo typeInfo, int typeIndex)
    {
        IntPtr typeAttributes = IntPtr.Zero;
        try
        {
            typeInfo.GetTypeAttr(out typeAttributes);
            var attributes = Marshal.PtrToStructure<TYPEATTR>(typeAttributes);
            GetDocumentation(typeInfo, -1, out var name, out _);
            var references = ReadImplementedReferences(typeInfo, attributes.cImplTypes);
            var members = ImmutableArray.CreateBuilder<TypeMemberObservation>(attributes.cFuncs + attributes.cVars);
            for (var index = 0; index < attributes.cFuncs; index++)
                members.Add(ReadFunction(typeInfo, index));
            for (var index = 0; index < attributes.cVars; index++)
                members.Add(ReadVariable(typeInfo, index, attributes.cFuncs));
            return new TypeObservation(
                new TypeIdentity(attributes.guid, name),
                ConvertTypeKind(attributes.typekind),
                (int)attributes.wTypeFlags,
                typeIndex,
                references,
                members.ToImmutable(),
                ImmutableDictionary<string, string>.Empty);
        }
        finally
        {
            _releaser.ReleaseTypeAttributes(typeInfo, typeAttributes);
        }
    }

    private ImmutableArray<TypeReferenceObservation> ReadImplementedReferences(ITypeInfo typeInfo, short count)
    {
        var references = ImmutableArray.CreateBuilder<TypeReferenceObservation>(count);
        for (var index = 0; index < count; index++)
        {
            try
            {
                typeInfo.GetRefTypeOfImplType(index, out var href);
                typeInfo.GetRefTypeInfo(href, out var referencedInfo);
                if (referencedInfo is null) continue;
                try
                {
                    IntPtr attributesPointer = IntPtr.Zero;
                    try
                    {
                        referencedInfo.GetTypeAttr(out attributesPointer);
                        var attributes = Marshal.PtrToStructure<TYPEATTR>(attributesPointer);
                        GetDocumentation(referencedInfo, -1, out var name, out _);
                        references.Add(new TypeReferenceObservation(
                            new TypeLibReference(attributes.guid, null, null, name),
                            $"implemented-interface[{index}]"));
                    }
                    finally
                    {
                        _releaser.ReleaseTypeAttributes(referencedInfo, attributesPointer);
                    }
                }
                finally
                {
                    _releaser.ReleaseComObject(referencedInfo);
                }
            }
            catch (COMException ex)
            {
                throw new TypeLibDescriptorException($"Unable to resolve implemented typelib reference {index}.", ex);
            }
        }

        return references.ToImmutable();
    }

    private TypeMemberObservation ReadFunction(ITypeInfo typeInfo, int index)
    {
        IntPtr descriptor = IntPtr.Zero;
        try
        {
            typeInfo.GetFuncDesc(index, out descriptor);
            var function = Marshal.PtrToStructure<FUNCDESC>(descriptor);
            GetDocumentation(typeInfo, function.memid, out var name, out _);
            var references = ReadSignatureReferences(typeInfo, function.memid);
            return new TypeMemberObservation(name, index, (int)function.invkind, function.elemdescFunc.tdesc.vt, function.wFuncFlags, function.memid, references, ImmutableDictionary<string, string>.Empty);
        }
        finally
        {
            _releaser.ReleaseFunctionDescriptor(typeInfo, descriptor);
        }
    }

    private TypeMemberObservation ReadVariable(ITypeInfo typeInfo, int index, int functionCount)
    {
        IntPtr descriptor = IntPtr.Zero;
        try
        {
            typeInfo.GetVarDesc(index, out descriptor);
            var variable = Marshal.PtrToStructure<VARDESC>(descriptor);
            GetDocumentation(typeInfo, variable.memid, out var name, out _);
            return new TypeMemberObservation(name, functionCount + index, 0, variable.elemdescVar.tdesc.vt, variable.wVarFlags, variable.memid, ImmutableArray<TypeReferenceObservation>.Empty, ImmutableDictionary<string, string>.Empty);
        }
        finally
        {
            _releaser.ReleaseVariableDescriptor(typeInfo, descriptor);
        }
    }

    private static ImmutableArray<TypeReferenceObservation> ReadSignatureReferences(ITypeInfo typeInfo, int memberId)
    {
        // GetDocumentation is intentionally the only portable member-level query here. Signature
        // TYPEDESC trees are native pointer graphs; the descriptor itself is still released by the caller.
        return ImmutableArray<TypeReferenceObservation>.Empty;
    }

    private static void GetDocumentation(ITypeLib typeLib, int memberId, out string name, out string documentation)
    {
        typeLib.GetDocumentation(memberId, out name, out documentation, out _, out _);
        name ??= string.Empty;
        documentation ??= string.Empty;
    }

    private static void GetDocumentation(ITypeInfo typeInfo, int memberId, out string name, out string documentation)
    {
        typeInfo.GetDocumentation(memberId, out name, out documentation, out _, out _);
        name ??= string.Empty;
        documentation ??= string.Empty;
    }

    private static TypeKind ConvertTypeKind(TYPEKIND kind) => kind switch
    {
        TYPEKIND.TKIND_ENUM => TypeKind.Enum,
        TYPEKIND.TKIND_RECORD => TypeKind.Record,
        TYPEKIND.TKIND_MODULE => TypeKind.Module,
        TYPEKIND.TKIND_INTERFACE => TypeKind.Interface,
        TYPEKIND.TKIND_DISPATCH => TypeKind.Dispatch,
        TYPEKIND.TKIND_COCLASS => TypeKind.CoClass,
        TYPEKIND.TKIND_ALIAS => TypeKind.Alias,
        TYPEKIND.TKIND_UNION => TypeKind.Union,
        _ => TypeKind.Unknown
    };
}
