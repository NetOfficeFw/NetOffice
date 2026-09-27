using System.Collections.Immutable;
using System.Globalization;

namespace NetOffice.CodeGen.TypeLib;

/// <summary>Schema identifier for raw observations consumed by this component.</summary>
public static class TypeLibObservationSchema
{
    public const string Version = "typelib-observation/v1";
}

public enum TypeLibSourceKind { File, Registry, Fixture }
public enum TypeKind { Enum = 0, Record = 1, Module = 2, Interface = 3, Dispatch = 4, CoClass = 5, Alias = 6, Union = 7, Max = 8, Unknown = 255 }

/// <summary>Stable identity of a typelib version.</summary>
public sealed record TypeLibIdentity(Guid LibraryId, ushort MajorVersion, ushort MinorVersion) : IComparable<TypeLibIdentity>
{
    public int CompareTo(TypeLibIdentity? other)
    {
        if (other is null) return 1;
        var result = LibraryId.CompareTo(other.LibraryId);
        if (result != 0) return result;
        result = MajorVersion.CompareTo(other.MajorVersion);
        return result != 0 ? result : MinorVersion.CompareTo(other.MinorVersion);
    }
    public override string ToString() => string.Create(CultureInfo.InvariantCulture, $"{LibraryId:D}/{MajorVersion}.{MinorVersion}");
}

public sealed record TypeIdentity(Guid TypeId, string Name);

/// <summary>Reference to another typelib, retaining enough information to diagnose omissions.</summary>
public sealed record TypeLibReference(Guid LibraryId, ushort? MajorVersion, ushort? MinorVersion, string? Name)
{
    public TypeLibIdentity? ExactIdentity => MajorVersion.HasValue && MinorVersion.HasValue ? new TypeLibIdentity(LibraryId, MajorVersion.Value, MinorVersion.Value) : null;
}

public sealed record TypeReferenceObservation(TypeLibReference Library, string Context);

/// <summary>Raw member observation. Values are deliberately unprojected and may be absent.</summary>
public sealed record TypeMemberObservation(
    string Name,
    int MemberIndex,
    int InvocationKind,
    short Vartype,
    short Flags,
    int? DispId,
    ImmutableArray<TypeReferenceObservation> References,
    ImmutableDictionary<string, string> Metadata);

/// <summary>Raw type descriptor observation.</summary>
public sealed record TypeObservation(
    TypeIdentity Identity,
    TypeKind Kind,
    int Flags,
    int TypeIndex,
    ImmutableArray<TypeReferenceObservation> References,
    ImmutableArray<TypeMemberObservation> Members,
    ImmutableDictionary<string, string> Metadata);

/// <summary>Current Channel acquisition facts attached to an observation.</summary>
public sealed record CurrentChannelMetadata(
    string ChannelGuid,
    string Build,
    string Sku,
    string Locale,
    string Architecture)
{
    public const string RequiredChannelGuid = "492350f6-3a01-4f97-b9c0-c7c6ddf67d60";
}

/// <summary>Acquisition and provenance facts. The optional timestamp is excluded from canonical hashes.</summary>
public sealed record TypeLibProvenance(
    TypeLibSourceKind SourceKind,
    string SourcePath,
    string BinarySha256,
    string? ImporterVersion,
    CurrentChannelMetadata? CurrentChannel,
    DateTimeOffset? CapturedAtUtc);

/// <summary>Complete immutable raw observation of one typelib.</summary>
public sealed record TypeLibObservation(
    string SchemaVersion,
    TypeLibIdentity Identity,
    string Name,
    int Lcid,
    int SysKind,
    TypeLibProvenance Provenance,
    ImmutableArray<TypeObservation> Types,
    ImmutableArray<TypeLibReference> References,
    ImmutableDictionary<string, string> Metadata)
{
    public bool IsKnownSchema => string.Equals(SchemaVersion, TypeLibObservationSchema.Version, StringComparison.Ordinal);
}
