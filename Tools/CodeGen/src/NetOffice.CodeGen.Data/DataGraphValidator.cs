using System.Text.RegularExpressions;

namespace NetOffice.CodeGen.Data;

public static class DataGraphValidator
{
    private static readonly Regex Sha256 = new("^[0-9a-fA-F]{64}$", RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static ValidationResult Validate(DataGraph graph, string? expectedDigest = null, string expectedSchema = DataSchema.Version)
    {
        ArgumentNullException.ThrowIfNull(graph);
        var issues = new List<ValidationIssue>();

        if (!string.Equals(graph.SchemaVersion, expectedSchema, StringComparison.Ordinal))
            issues.Add(new("schema.version", $"Expected {expectedSchema}, got {graph.SchemaVersion}.", "schemaVersion"));
        if (!string.Equals(graph.Serialization, DataSchema.Serialization, StringComparison.Ordinal))
            issues.Add(new("schema.serialization", $"Expected {DataSchema.Serialization}, got {graph.Serialization}.", "serialization"));
        if (graph.Source is null)
            issues.Add(new("source.required", "A graph source is required.", "source"));
        else
        {
            ValidateSha(graph.Source.Sha256, issues, "source.sha256");
            if (string.IsNullOrWhiteSpace(graph.Source.Kind))
                issues.Add(new("source.kind", "A source kind is required.", "source.kind"));
        }

        ValidateUnique(graph.Libraries.Select(static item => item.LogicalId), issues, "libraries");
        ValidateUnique(graph.Types.Select(static item => item.LogicalId), issues, "types");
        ValidateUnique(graph.Members.Select(static item => item.LogicalId), issues, "members");
        ValidateUnique(graph.AccessorGroups.Select(static item => item.LogicalId), issues, "accessorGroups");
        ValidateUnique(graph.Ambiguities.Select(static item => item.LogicalId), issues, "ambiguities");

        var libraryIds = graph.Libraries.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var groups = graph.AccessorGroups
            .GroupBy(static item => item.LogicalId, StringComparer.Ordinal)
            .ToDictionary(static group => group.Key, static group => group.First(), StringComparer.Ordinal);
        var typeIds = graph.Types.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var memberIds = graph.Members.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);

        foreach (var library in graph.Libraries)
        {
            RequireText(library.LogicalId, issues, "libraries.logicalId");
            RequireText(library.Name, issues, $"libraries[{library.LogicalId}].name");
            if (!Guid.TryParse(library.Guid, out _))
                issues.Add(new("library.guid", "Library GUID is not valid.", $"libraries[{library.LogicalId}].guid"));
            ValidateProvenance(library.Provenance, issues, $"libraries[{library.LogicalId}].provenance");
        }

        foreach (var type in graph.Types)
        {
            RequireText(type.LogicalId, issues, "types.logicalId");
            RequireText(type.Name, issues, $"types[{type.LogicalId}].name");
            RequireReference(type.LibraryId, libraryIds, issues, $"types[{type.LogicalId}].libraryId");
            foreach (var baseTypeId in type.BaseTypeIds)
                RequireReference(baseTypeId, typeIds, issues, $"types[{type.LogicalId}].baseTypeIds");
            ValidateProvenance(type.Provenance, issues, $"types[{type.LogicalId}].provenance");
        }

        foreach (var member in graph.Members)
        {
            RequireText(member.LogicalId, issues, "members.logicalId");
            RequireText(member.Name, issues, $"members[{member.LogicalId}].name");
            RequireReference(member.TypeId, typeIds, issues, $"members[{member.LogicalId}].typeId");
            if (member.AccessorGroupId is not null)
                RequireReference(member.AccessorGroupId, groups.Keys, issues, $"members[{member.LogicalId}].accessorGroupId");
            ValidateProvenance(member.Provenance, issues, $"members[{member.LogicalId}].provenance");
        }

        foreach (var group in graph.AccessorGroups)
        {
            RequireText(group.LogicalId, issues, "accessorGroups.logicalId");
            RequireText(group.Name, issues, $"accessorGroups[{group.LogicalId}].name");
            RequireReference(group.TypeId, typeIds, issues, $"accessorGroups[{group.LogicalId}].typeId");
            foreach (var memberId in group.MemberIds)
            {
                RequireReference(memberId, memberIds, issues, $"accessorGroups[{group.LogicalId}].memberIds");
                var member = graph.Members.FirstOrDefault(item => item.LogicalId == memberId);
                if (member is not null && !string.Equals(member.AccessorGroupId, group.LogicalId, StringComparison.Ordinal))
                    issues.Add(new("accessor.member-mismatch", "Member and accessor group disagree.", $"accessorGroups[{group.LogicalId}].memberIds"));
            }
        }

        foreach (var ambiguity in graph.Ambiguities)
        {
            RequireText(ambiguity.LogicalId, issues, "ambiguities.logicalId");
            RequireText(ambiguity.Kind, issues, $"ambiguities[{ambiguity.LogicalId}].kind");
            RequireText(ambiguity.Reason, issues, $"ambiguities[{ambiguity.LogicalId}].reason");
            ValidateProvenance(ambiguity.Provenance, issues, $"ambiguities[{ambiguity.LogicalId}].provenance");
        }

        var actualDigest = CanonicalJson.ComputeDigest(graph);
        if (string.IsNullOrWhiteSpace(graph.Digest) || !Sha256.IsMatch(graph.Digest))
            issues.Add(new("digest.format", "Graph digest must be a SHA-256 hex string.", "digest"));
        else if (!string.Equals(graph.Digest, actualDigest, StringComparison.OrdinalIgnoreCase))
            issues.Add(new("digest.mismatch", $"Expected {actualDigest}, got {graph.Digest}.", "digest"));
        if (expectedDigest is not null && !string.Equals(expectedDigest, actualDigest, StringComparison.OrdinalIgnoreCase))
            issues.Add(new("digest.expected-mismatch", $"Expected pinned digest {expectedDigest}, got {actualDigest}.", "digest"));

        return new ValidationResult(issues);
    }

    public static void ValidateOrThrow(DataGraph graph, string? expectedDigest = null, string expectedSchema = DataSchema.Version)
        => Validate(graph, expectedDigest, expectedSchema).ThrowIfInvalid();

    private static void ValidateProvenance(Provenance provenance, ICollection<ValidationIssue> issues, string path)
    {
        RequireText(provenance.SourcePath, issues, $"{path}.sourcePath");
        ValidateSha(provenance.SourceSha256, issues, $"{path}.sourceSha256");
    }

    private static void ValidateSha(string value, ICollection<ValidationIssue> issues, string path)
    {
        if (!Sha256.IsMatch(value ?? ""))
            issues.Add(new("sha256.format", "SHA-256 must contain exactly 64 hexadecimal characters.", path));
    }

    private static void ValidateUnique(IEnumerable<string> values, ICollection<ValidationIssue> issues, string path)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var value in values)
            if (!seen.Add(value))
                issues.Add(new("id.duplicate", $"Duplicate logical ID {value}.", path));
    }

    private static void RequireReference(string value, IEnumerable<string> known, ICollection<ValidationIssue> issues, string path)
    {
        RequireText(value, issues, path);
        if (!known.Contains(value, StringComparer.Ordinal))
            issues.Add(new("reference.missing", $"Unknown logical ID {value}.", path));
    }

    private static void RequireText(string value, ICollection<ValidationIssue> issues, string path)
    {
        if (string.IsNullOrWhiteSpace(value))
            issues.Add(new("value.required", "A non-empty value is required.", path));
    }
}
