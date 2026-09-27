using System.Collections.Immutable;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.Json;

namespace NetOffice.CodeGen.TypeLib;

/// <summary>Canonical, versioned JSON encoding for raw observations.</summary>
public static class TypeLibObservationSerializer
{
    public static byte[] Serialize(TypeLibObservation observation)
    {
        ArgumentNullException.ThrowIfNull(observation);
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false }))
        {
            WriteObservation(writer, observation);
            writer.Flush();
        }

        return stream.ToArray();
    }

    public static string SerializeToString(TypeLibObservation observation) => Encoding.UTF8.GetString(Serialize(observation));

    public static TypeLibObservation Deserialize(ReadOnlySpan<byte> json)
    {
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };
        options.Converters.Add(new JsonStringEnumConverter());
        var value = JsonSerializer.Deserialize<TypeLibObservation>(json, options);
        return value ?? throw new JsonException("Typelib observation JSON was null.");
    }

    public static string Sha256(TypeLibObservation observation)
    {
        var hash = SHA256.HashData(Serialize(observation));
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static void WriteObservation(Utf8JsonWriter writer, TypeLibObservation observation)
    {
        writer.WriteStartObject();
        writer.WriteString("schemaVersion", observation.SchemaVersion);
        writer.WritePropertyName("identity");
        writer.WriteStartObject();
        writer.WriteString("libraryId", observation.Identity.LibraryId);
        writer.WriteNumber("majorVersion", observation.Identity.MajorVersion);
        writer.WriteNumber("minorVersion", observation.Identity.MinorVersion);
        writer.WriteEndObject();
        writer.WriteString("name", observation.Name);
        writer.WriteNumber("lcid", observation.Lcid);
        writer.WriteNumber("sysKind", observation.SysKind);
        writer.WritePropertyName("provenance");
        WriteProvenance(writer, observation.Provenance);
        writer.WritePropertyName("references");
        writer.WriteStartArray();
        foreach (var reference in observation.References.OrderBy(r => r.LibraryId).ThenBy(r => r.MajorVersion).ThenBy(r => r.MinorVersion).ThenBy(r => r.Name, StringComparer.Ordinal))
            WriteReference(writer, reference);
        writer.WriteEndArray();
        writer.WritePropertyName("types");
        writer.WriteStartArray();
        foreach (var type in observation.Types.OrderBy(t => t.Identity.TypeId).ThenBy(t => t.Identity.Name, StringComparer.Ordinal))
            WriteType(writer, type);
        writer.WriteEndArray();
        WriteMetadata(writer, observation.Metadata);
        writer.WriteEndObject();
    }

    private static void WriteProvenance(Utf8JsonWriter writer, TypeLibProvenance provenance)
    {
        writer.WriteStartObject();
        writer.WriteString("sourceKind", provenance.SourceKind.ToString());
        writer.WriteString("sourcePath", provenance.SourcePath);
        writer.WriteString("binarySha256", provenance.BinarySha256.ToLowerInvariant());
        if (provenance.ImporterVersion is null) writer.WriteNull("importerVersion"); else writer.WriteString("importerVersion", provenance.ImporterVersion);
        writer.WritePropertyName("currentChannel");
        if (provenance.CurrentChannel is null) writer.WriteNullValue(); else WriteChannel(writer, provenance.CurrentChannel);
        writer.WriteEndObject();
    }

    private static void WriteChannel(Utf8JsonWriter writer, CurrentChannelMetadata channel)
    {
        writer.WriteStartObject();
        writer.WriteString("channelGuid", channel.ChannelGuid.ToLowerInvariant());
        writer.WriteString("build", channel.Build);
        writer.WriteString("sku", channel.Sku);
        writer.WriteString("locale", channel.Locale);
        writer.WriteString("architecture", channel.Architecture);
        writer.WriteEndObject();
    }

    private static void WriteType(Utf8JsonWriter writer, TypeObservation type)
    {
        writer.WriteStartObject();
        writer.WritePropertyName("identity");
        writer.WriteStartObject();
        writer.WriteString("typeId", type.Identity.TypeId);
        writer.WriteString("name", type.Identity.Name);
        writer.WriteEndObject();
        writer.WriteString("kind", type.Kind.ToString());
        writer.WriteNumber("flags", type.Flags);
        writer.WriteNumber("typeIndex", type.TypeIndex);
        writer.WritePropertyName("references");
        writer.WriteStartArray();
        foreach (var reference in type.References.OrderBy(r => r.Library.LibraryId).ThenBy(r => r.Library.MajorVersion).ThenBy(r => r.Library.MinorVersion).ThenBy(r => r.Library.Name, StringComparer.Ordinal).ThenBy(r => r.Context, StringComparer.Ordinal))
        {
            writer.WriteStartObject();
            writer.WritePropertyName("library");
            WriteReference(writer, reference.Library);
            writer.WriteString("context", reference.Context);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WritePropertyName("members");
        writer.WriteStartArray();
        foreach (var member in type.Members.OrderBy(m => m.MemberIndex).ThenBy(m => m.Name, StringComparer.Ordinal))
        {
            writer.WriteStartObject();
            writer.WriteString("name", member.Name);
            writer.WriteNumber("memberIndex", member.MemberIndex);
            writer.WriteNumber("invocationKind", member.InvocationKind);
            writer.WriteNumber("vartype", member.Vartype);
            writer.WriteNumber("flags", member.Flags);
            if (member.DispId.HasValue) writer.WriteNumber("dispId", member.DispId.Value); else writer.WriteNull("dispId");
            writer.WritePropertyName("references");
            writer.WriteStartArray();
            foreach (var reference in member.References.OrderBy(r => r.Library.LibraryId).ThenBy(r => r.Library.MajorVersion).ThenBy(r => r.Library.MinorVersion).ThenBy(r => r.Library.Name, StringComparer.Ordinal).ThenBy(r => r.Context, StringComparer.Ordinal))
            {
                writer.WriteStartObject();
                writer.WritePropertyName("library");
                WriteReference(writer, reference.Library);
                writer.WriteString("context", reference.Context);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            WriteMetadata(writer, member.Metadata);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        WriteMetadata(writer, type.Metadata);
        writer.WriteEndObject();
    }

    private static void WriteReference(Utf8JsonWriter writer, TypeLibReference reference)
    {
        writer.WriteStartObject();
        writer.WriteString("libraryId", reference.LibraryId);
        if (reference.MajorVersion.HasValue) writer.WriteNumber("majorVersion", reference.MajorVersion.Value); else writer.WriteNull("majorVersion");
        if (reference.MinorVersion.HasValue) writer.WriteNumber("minorVersion", reference.MinorVersion.Value); else writer.WriteNull("minorVersion");
        if (reference.Name is null) writer.WriteNull("name"); else writer.WriteString("name", reference.Name);
        writer.WriteEndObject();
    }

    private static void WriteMetadata(Utf8JsonWriter writer, ImmutableDictionary<string, string> metadata)
    {
        writer.WritePropertyName("metadata");
        writer.WriteStartObject();
        foreach (var item in metadata.OrderBy(x => x.Key, StringComparer.Ordinal))
            writer.WriteString(item.Key, item.Value);
        writer.WriteEndObject();
    }
}
