using System.Collections.Immutable;
using System.Text;
using System.Text.Json;

namespace NetOffice.CodeGen.TypeLib;

/// <summary>A reference that could not be resolved to an imported typelib.</summary>
public sealed record UnresolvedTypeLibReference(
    TypeLibIdentity Owner,
    TypeLibReference Reference,
    string Context,
    string Reason);

/// <summary>One deterministic dependency node.</summary>
public sealed record TypeLibDependencyNode(TypeLibIdentity Identity, ImmutableArray<TypeLibIdentity> Dependencies);

/// <summary>Validated graph and explicit unresolved-reference diagnostics.</summary>
public sealed class TypeLibDependencyGraph
{
    private TypeLibDependencyGraph(
        ImmutableArray<TypeLibDependencyNode> nodes,
        ImmutableArray<UnresolvedTypeLibReference> unresolved,
        ImmutableArray<string> errors)
    {
        Nodes = nodes;
        UnresolvedReferences = unresolved;
        Errors = errors;
    }

    public ImmutableArray<TypeLibDependencyNode> Nodes { get; }
    public ImmutableArray<UnresolvedTypeLibReference> UnresolvedReferences { get; }
    public ImmutableArray<string> Errors { get; }
    public bool IsValid => Errors.IsEmpty && UnresolvedReferences.IsEmpty;

    public static TypeLibDependencyGraph Build(IEnumerable<TypeLibObservation> observations)
    {
        ArgumentNullException.ThrowIfNull(observations);
        var ordered = observations.OrderBy(x => x.Identity).ToArray();
        var errors = ImmutableArray.CreateBuilder<string>();
        var duplicate = ordered.GroupBy(x => x.Identity).Where(x => x.Count() > 1);
        foreach (var group in duplicate)
            errors.Add($"Duplicate typelib identity '{group.Key}'.");

        var known = ordered.Select(x => x.Identity).ToHashSet();
        var byGuid = ordered.GroupBy(x => x.Identity.LibraryId).ToDictionary(x => x.Key, x => x.Select(y => y.Identity).OrderBy(y => y).ToArray());
        var unresolved = ImmutableArray.CreateBuilder<UnresolvedTypeLibReference>();
        var nodes = ImmutableArray.CreateBuilder<TypeLibDependencyNode>();
        foreach (var observation in ordered.GroupBy(x => x.Identity).Select(x => x.First()).OrderBy(x => x.Identity))
        {
            var refs = observation.References
                .Select(reference => (reference: reference, context: "typelib"))
                .Concat(observation.Types.SelectMany(type => type.References.Select(reference => (reference: reference.Library, context: reference.Context))))
                .Concat(observation.Types.SelectMany(type => type.Members.SelectMany(member => member.References.Select(reference => (reference: reference.Library, context: reference.Context)))))
                .ToArray();
            var dependencies = new HashSet<TypeLibIdentity>();
            foreach (var item in refs.OrderBy(x => x.reference.LibraryId).ThenBy(x => x.reference.MajorVersion).ThenBy(x => x.reference.MinorVersion).ThenBy(x => x.context, StringComparer.Ordinal))
            {
                var reference = item.reference;
                var context = item.context;
                var exact = reference.ExactIdentity;
                if (exact is not null && known.Contains(exact))
                {
                    dependencies.Add(exact);
                    continue;
                }

                var candidates = byGuid.TryGetValue(reference.LibraryId, out var values) ? values : Array.Empty<TypeLibIdentity>();
                if (exact is null && candidates.Length == 1)
                {
                    dependencies.Add(candidates[0]);
                    continue;
                }

                var reason = exact is not null
                    ? $"No imported typelib matches {exact}."
                    : candidates.Length == 0
                        ? "No imported typelib has the referenced library GUID."
                        : "Reference omits a version and matches multiple imported versions.";
                unresolved.Add(new UnresolvedTypeLibReference(observation.Identity, reference, context, reason));
            }

            nodes.Add(new TypeLibDependencyNode(observation.Identity, dependencies.OrderBy(x => x).ToImmutableArray()));
        }

        return new TypeLibDependencyGraph(nodes.ToImmutable(), unresolved.ToImmutable(), errors.ToImmutable());
    }
}

/// <summary>Conflict kind emitted by deterministic merge.</summary>
public enum TypeLibMergeConflictKind
{
    DuplicateIdentity,
    StaleInput,
    UnknownSchema,
    InvalidObservation
}

/// <summary>One merge conflict with stable hashes and a machine-readable reason.</summary>
public sealed record TypeLibMergeConflict(
    TypeLibMergeConflictKind Kind,
    TypeLibIdentity? Identity,
    string? ExistingHash,
    string? IncomingHash,
    string Reason);

/// <summary>Absence is never silently treated as deletion.</summary>
public sealed record TypeLibRemovalCandidate(TypeLibIdentity Identity, string LastObservationHash, string Reason);

/// <summary>Optional compare-and-swap expectations for rejecting stale imports.</summary>
public sealed record TypeLibMergeRequest(
    IReadOnlyDictionary<TypeLibIdentity, string>? ExpectedBaseHashes,
    bool RejectUnknownSchema = true);

/// <summary>Complete deterministic merge outcome.</summary>
public sealed class TypeLibMergeResult
{
    internal TypeLibMergeResult(
        ImmutableArray<TypeLibObservation> observations,
        ImmutableArray<TypeLibMergeConflict> conflicts,
        ImmutableArray<TypeLibRemovalCandidate> removalCandidates,
        ImmutableArray<UnresolvedTypeLibReference> unresolvedReferences)
    {
        Observations = observations;
        Conflicts = conflicts;
        RemovalCandidates = removalCandidates;
        UnresolvedReferences = unresolvedReferences;
    }

    public ImmutableArray<TypeLibObservation> Observations { get; }
    public ImmutableArray<TypeLibMergeConflict> Conflicts { get; }
    public ImmutableArray<TypeLibRemovalCandidate> RemovalCandidates { get; }
    public ImmutableArray<UnresolvedTypeLibReference> UnresolvedReferences { get; }
    public bool IsSuccess => Conflicts.IsEmpty && UnresolvedReferences.IsEmpty;

    public byte[] SerializeReport()
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream))
        {
            writer.WriteStartObject();
            writer.WritePropertyName("conflicts");
            writer.WriteStartArray();
            foreach (var conflict in Conflicts.OrderBy(x => x.Kind).ThenBy(x => x.Identity).ThenBy(x => x.ExistingHash, StringComparer.Ordinal).ThenBy(x => x.IncomingHash, StringComparer.Ordinal).ThenBy(x => x.Reason, StringComparer.Ordinal))
            {
                writer.WriteStartObject();
                writer.WriteString("kind", conflict.Kind.ToString());
                if (conflict.Identity is null) writer.WriteNull("identity"); else writer.WriteString("identity", conflict.Identity.ToString());
                if (conflict.ExistingHash is null) writer.WriteNull("existingHash"); else writer.WriteString("existingHash", conflict.ExistingHash);
                if (conflict.IncomingHash is null) writer.WriteNull("incomingHash"); else writer.WriteString("incomingHash", conflict.IncomingHash);
                writer.WriteString("reason", conflict.Reason);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WritePropertyName("removalCandidates");
            writer.WriteStartArray();
            foreach (var candidate in RemovalCandidates.OrderBy(x => x.Identity))
            {
                writer.WriteStartObject();
                writer.WriteString("identity", candidate.Identity.ToString());
                writer.WriteString("lastObservationHash", candidate.LastObservationHash);
                writer.WriteString("reason", candidate.Reason);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WritePropertyName("unresolvedReferences");
            writer.WriteStartArray();
            foreach (var unresolved in UnresolvedReferences.OrderBy(x => x.Owner).ThenBy(x => x.Reference.LibraryId).ThenBy(x => x.Reference.MajorVersion).ThenBy(x => x.Reference.MinorVersion).ThenBy(x => x.Reference.Name, StringComparer.Ordinal).ThenBy(x => x.Context, StringComparer.Ordinal))
            {
                writer.WriteStartObject();
                writer.WriteString("owner", unresolved.Owner.ToString());
                writer.WriteString("libraryId", unresolved.Reference.LibraryId);
                if (unresolved.Reference.MajorVersion.HasValue) writer.WriteNumber("majorVersion", unresolved.Reference.MajorVersion.Value); else writer.WriteNull("majorVersion");
                if (unresolved.Reference.MinorVersion.HasValue) writer.WriteNumber("minorVersion", unresolved.Reference.MinorVersion.Value); else writer.WriteNull("minorVersion");
                if (unresolved.Reference.Name is null) writer.WriteNull("name"); else writer.WriteString("name", unresolved.Reference.Name);
                writer.WriteString("context", unresolved.Context);
                writer.WriteString("reason", unresolved.Reason);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return stream.ToArray();
    }

    public string SerializeReportText() => Encoding.UTF8.GetString(SerializeReport());
}

/// <summary>Merges immutable observations without nondeterministic overwrite or deletion.</summary>
public static class TypeLibMerger
{
    public static TypeLibMergeResult Merge(
        IEnumerable<TypeLibObservation> existing,
        IEnumerable<TypeLibObservation> incoming,
        TypeLibMergeRequest? request = null)
    {
        ArgumentNullException.ThrowIfNull(existing);
        ArgumentNullException.ThrowIfNull(incoming);
        request ??= new TypeLibMergeRequest(null);
        var conflicts = ImmutableArray.CreateBuilder<TypeLibMergeConflict>();
        var existingByIdentity = existing.OrderBy(x => x.Identity).GroupBy(x => x.Identity).ToArray();
        foreach (var group in existingByIdentity.Where(x => x.Count() > 1))
        {
            var hashes = group.Select(TypeLibObservationSerializer.Sha256).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.Ordinal).ToArray();
            if (hashes.Length > 1)
                conflicts.Add(new TypeLibMergeConflict(TypeLibMergeConflictKind.DuplicateIdentity, group.Key, hashes[0], hashes[^1], "Existing input contains divergent observations with the same identity."));
        }
        foreach (var observation in existingByIdentity.Select(x => x.First()))
        {
            if (request.RejectUnknownSchema && !observation.IsKnownSchema)
                conflicts.Add(new TypeLibMergeConflict(TypeLibMergeConflictKind.UnknownSchema, observation.Identity, null, TypeLibObservationSerializer.Sha256(observation), $"Unsupported observation schema '{observation.SchemaVersion}'."));
        }
        var current = existingByIdentity.ToDictionary(x => x.Key, x => x.First());
        var incomingByIdentity = incoming.OrderBy(x => x.Identity).ToArray();
        foreach (var observation in incomingByIdentity)
        {
            if (!string.Equals(observation.SchemaVersion, TypeLibObservationSchema.Version, StringComparison.Ordinal))
            {
                if (request.RejectUnknownSchema)
                    conflicts.Add(new TypeLibMergeConflict(TypeLibMergeConflictKind.UnknownSchema, observation.Identity, null, null, $"Unsupported observation schema '{observation.SchemaVersion}'."));
                continue;
            }

            var incomingHash = TypeLibObservationSerializer.Sha256(observation);
            if (request.ExpectedBaseHashes is not null && request.ExpectedBaseHashes.TryGetValue(observation.Identity, out var expected) && current.TryGetValue(observation.Identity, out var actualBase))
            {
                var actual = TypeLibObservationSerializer.Sha256(actualBase);
                if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                {
                    conflicts.Add(new TypeLibMergeConflict(TypeLibMergeConflictKind.StaleInput, observation.Identity, actual, incomingHash, $"Expected base hash '{expected}' but found '{actual}'."));
                    continue;
                }
            }

            if (current.TryGetValue(observation.Identity, out var previous))
            {
                var previousHash = TypeLibObservationSerializer.Sha256(previous);
                if (!string.Equals(previousHash, incomingHash, StringComparison.OrdinalIgnoreCase))
                    conflicts.Add(new TypeLibMergeConflict(TypeLibMergeConflictKind.DuplicateIdentity, observation.Identity, previousHash, incomingHash, "Two observations have the same identity but different content."));
                continue;
            }
            current.Add(observation.Identity, observation);
        }

        var incomingIdentities = incomingByIdentity.Select(x => x.Identity).ToHashSet();
        var removals = existing
            .Where(x => !incomingIdentities.Contains(x.Identity))
            .OrderBy(x => x.Identity)
            .Select(x => new TypeLibRemovalCandidate(x.Identity, TypeLibObservationSerializer.Sha256(x), "Absent from import; maintainer approval is required before removal."))
            .ToImmutableArray();
        var observations = current.Values.OrderBy(x => x.Identity).ToImmutableArray();
        var graph = TypeLibDependencyGraph.Build(observations);
        return new TypeLibMergeResult(observations, conflicts.ToImmutable(), removals, graph.UnresolvedReferences);
    }
}
