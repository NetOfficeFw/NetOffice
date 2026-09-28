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

        ValidateUnique(graph.Projects.Select(static item => item.LogicalId), issues, "projects");
        ValidateUnique(graph.Libraries.Select(static item => item.LogicalId), issues, "libraries");
        ValidateUnique(graph.AbsentFacts.Select(static item => item.LogicalId), issues, "absentFacts");
        ValidateUnique(graph.Types.Select(static item => item.LogicalId), issues, "types");
        ValidateUnique(graph.Members.Select(static item => item.LogicalId), issues, "members");
        ValidateUnique(graph.Values.Select(static item => item.LogicalId), issues, "values");
        ValidateUnique(graph.AccessorGroups.Select(static item => item.LogicalId), issues, "accessorGroups");
        ValidateUnique(graph.SupportObservations.Select(static item => item.LogicalId), issues, "supportObservations");
        ValidateUnique(graph.Aliases.Select(static item => item.LogicalId), issues, "aliases");
        ValidateUnique(graph.Unifications.Select(static item => item.LogicalId), issues, "unifications");
        ValidateUnique(graph.InvocationEvidence.Select(static item => item.LogicalId), issues, "invocationEvidence");
        ValidateUnique(graph.AccessorEvidence.Select(static item => item.LogicalId), issues, "accessorEvidence");
        ValidateUnique(graph.Unknowns.Select(static item => item.LogicalId), issues, "unknowns");
        ValidateUnique(graph.StaleRecords.Select(static item => item.LogicalId), issues, "staleRecords");
        ValidateUnique(graph.Ambiguities.Select(static item => item.LogicalId), issues, "ambiguities");

        var projectIds = graph.Projects.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var libraryIds = graph.Libraries.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var typeIds = graph.Types.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var memberIds = graph.Members.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var valueIds = graph.Values.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var groupIds = graph.AccessorGroups.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var observationIds = graph.SupportObservations.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var evidenceIds = graph.InvocationEvidence.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var accessorEvidenceIds = graph.AccessorEvidence.Select(static item => item.LogicalId).ToHashSet(StringComparer.Ordinal);
        var allIds = projectIds.Concat(libraryIds).Concat(typeIds).Concat(memberIds).Concat(valueIds).Concat(groupIds).ToHashSet(StringComparer.Ordinal);
        var membersById = graph.Members.GroupBy(static item => item.LogicalId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.First(), StringComparer.Ordinal);

        foreach (var project in graph.Projects)
        {
            var path = $"projects[{project.LogicalId}]";
            RequireText(project.LogicalId, issues, "projects.logicalId");
            RequireText(project.Name, issues, $"{path}.name");
            RequireText(project.Namespace, issues, $"{path}.namespace");
            RequireText(project.SourceKey, issues, $"{path}.sourceKey");
            RequireText(project.Version, issues, $"{path}.version");
            RequireText(project.FileVersion, issues, $"{path}.fileVersion");
            if (project.SourceCategories.Count == 0)
                issues.Add(new("project.source-categories", "A project must declare its included XML source categories.", $"{path}.sourceCategories"));
            foreach (var libraryId in project.LibraryIds)
                RequireReference(libraryId, libraryIds, issues, $"{path}.libraryIds");
            foreach (var projectId in project.ReferenceProjectIds)
                RequireReference(projectId, projectIds, issues, $"{path}.referenceProjectIds");
            ValidateProvenance(project.Provenance, issues, $"{path}.provenance");
        }

        foreach (var library in graph.Libraries)
        {
            var path = $"libraries[{library.LogicalId}]";
            RequireText(library.LogicalId, issues, "libraries.logicalId");
            RequireText(library.Name, issues, $"{path}.name");
            if (!Guid.TryParse(library.Guid, out _))
                issues.Add(new("library.guid", "Library GUID is not valid.", $"{path}.guid"));
            if (graph.Projects.Count != 0)
            {
                RequireText(library.SourceKey, issues, $"{path}.sourceKey");
                RequireText(library.Major, issues, $"{path}.major");
                RequireText(library.Minor, issues, $"{path}.minor");
            }
            foreach (var dependency in library.Dependencies)
            {
                RequireText(dependency.Name, issues, $"{path}.dependencies.name");
                if (!Guid.TryParse(dependency.Guid, out _))
                    issues.Add(new("library.dependency-guid", "Dependency GUID is not valid.", $"{path}.dependencies.guid"));
                RequireText(dependency.Major, issues, $"{path}.dependencies.major");
                RequireText(dependency.Minor, issues, $"{path}.dependencies.minor");
                foreach (var targetId in dependency.TargetLibraryIds)
                    RequireReference(targetId, libraryIds, issues, $"{path}.dependencies.targetLibraryIds");
                ValidateProvenance(dependency.Provenance, issues, $"{path}.dependencies.provenance");
            }
            ValidateProvenance(library.Provenance, issues, $"{path}.provenance");
        }
        foreach (var type in graph.Types)
        {
            var path = $"types[{type.LogicalId}]";
            RequireText(type.LogicalId, issues, "types.logicalId");
            RequireText(type.Name, issues, $"{path}.name");
            RequireText(type.Kind, issues, $"{path}.kind");
            RequireText(type.SourceKey, issues, $"{path}.sourceKey");
            RequireReference(type.LibraryId, libraryIds, issues, $"{path}.libraryId");
            if (graph.Projects.Count != 0)
            {
                RequireReference(type.ProjectId, projectIds, issues, $"{path}.projectId");
                RequireText(type.Namespace, issues, $"{path}.namespace");
                RequireText(type.SourceCategory, issues, $"{path}.sourceCategory");
            }
            if (type.DeclaredGuid is not null && !Guid.TryParse(type.DeclaredGuid, out _))
                issues.Add(new("type.guid", "Declared type GUID is not valid.", $"{path}.declaredGuid"));
            ValidateIdentifiers(type.GuidObservations, libraryIds, issues, $"{path}.guidObservations", requireGuid: true);
            foreach (var reference in type.ReferenceObservations)
            {
                RequireText(reference.Kind, issues, $"{path}.referenceObservations.kind");
                RequireReference(reference.TargetTypeId, typeIds, issues, $"{path}.referenceObservations.targetTypeId");
                if (reference.LibraryIds.Count == 0)
                    issues.Add(new("type-reference.libraries", "Type reference observation must identify at least one supporting library.", $"{path}.referenceObservations.libraryIds"));
                foreach (var libraryId in reference.LibraryIds)
                    RequireReference(libraryId, libraryIds, issues, $"{path}.referenceObservations.libraryIds");
                ValidateProvenance(reference.Provenance, issues, $"{path}.referenceObservations.provenance");
            }
            foreach (var baseTypeId in type.BaseTypeIds)
                RequireReference(baseTypeId, typeIds, issues, $"{path}.baseTypeIds");
            foreach (var interfaceId in type.DefaultInterfaceIds)
                RequireReference(interfaceId, typeIds, issues, $"{path}.defaultInterfaceIds");
            foreach (var interfaceId in type.EventInterfaceIds)
                RequireReference(interfaceId, typeIds, issues, $"{path}.eventInterfaceIds");
            foreach (var observationId in type.SupportObservationIds)
                RequireReference(observationId, observationIds, issues, $"{path}.supportObservationIds");
            ValidateProvenance(type.Provenance, issues, $"{path}.provenance");
        }
        foreach (var member in graph.Members)
        {
            var memberPath = $"members[{member.LogicalId}]";
            RequireText(member.LogicalId, issues, "members.logicalId");
            RequireText(member.Name, issues, $"{memberPath}.name");
            RequireText(member.Kind, issues, $"{memberPath}.kind");
            RequireText(member.SourceKey, issues, $"{memberPath}.sourceKey");
            RequireReference(member.TypeId, typeIds, issues, $"{memberPath}.typeId");
            ValidateIdentifiers(member.DispIdObservations, libraryIds, issues, $"{memberPath}.dispIdObservations", requireInteger: true);
            if (member.ReturnTypeReference is not null)
                ValidateTypeReference(member.ReturnTypeReference, typeIds, issues, $"{memberPath}.returnTypeReference");
            if (member.AccessorGroupId is not null)
                RequireReference(member.AccessorGroupId, groupIds, issues, $"{memberPath}.accessorGroupId");
            if (member.AccessorKind is not null && member.AccessorKind is not ("get" or "put" or "putref"))
                issues.Add(new("accessor.kind", "AccessorKind must be get, put, or putref.", $"{memberPath}.accessorKind"));
            if (member.AccessorKind is not null && member.AccessorGroupId is null)
                issues.Add(new("accessor.group-required", "AccessorKind requires an accessor group.", $"{memberPath}.accessorGroupId"));
            if (member.Parameters.Count != 0)
            {
                if (member.ParameterTypes.Count != member.Parameters.Count)
                    issues.Add(new("signature.parameter-types-mismatch", "ParameterTypes must match the emit-ready parameter list.", $"{memberPath}.parameters"));
                for (var index = 0; index < member.Parameters.Count; index++)
                {
                    var parameter = member.Parameters[index];
                    var parameterPath = $"{memberPath}.parameters[{index}]";
                    RequireText(parameter.Name, issues, $"{parameterPath}.name");
                    RequireText(parameter.Type, issues, $"{parameterPath}.type");
                    if (parameter.RefKind is not ("value" or "ref" or "out"))
                        issues.Add(new("signature.ref-kind", "RefKind must be value, ref, or out.", $"{parameterPath}.refKind"));
                    if (parameter.DefaultValue is not null && !parameter.IsOptional)
                        issues.Add(new("signature.default-without-optional", "A default value requires an optional parameter.", parameterPath));
                    ValidateProvenance(parameter.Provenance, issues, $"{parameterPath}.provenance");
                    if (parameter.HasDefaultValue && parameter.DefaultValue is null)
                        issues.Add(new("signature.default-value-missing", "HasDefaultValue requires the source default value, including an explicit empty string.", parameterPath));
                    if (!parameter.HasDefaultValue && parameter.DefaultValue is not null)
                        issues.Add(new("signature.default-value-unexpected", "DefaultValue requires HasDefaultValue.", parameterPath));
                    if (parameter.TypeReference is not null)
                        ValidateTypeReference(parameter.TypeReference, typeIds, issues, $"{parameterPath}.typeReference");
                }
            }
            else if (member.ParameterTypes.Count != 0)
                issues.Add(new("signature.parameters-missing", "Parameter facts are required when parameter types are present.", $"{memberPath}.parameters"));
            foreach (var observationId in member.SupportObservationIds)
                RequireReference(observationId, observationIds, issues, $"{memberPath}.supportObservationIds");
            foreach (var libraryId in member.SignatureLibraryIds)
                RequireReference(libraryId, libraryIds, issues, $"{memberPath}.signatureLibraryIds");
            if (member.InvocationEvidenceId is not null)
                RequireReference(member.InvocationEvidenceId, evidenceIds, issues, $"{memberPath}.invocationEvidenceId");
            ValidateProvenance(member.Provenance, issues, $"{memberPath}.provenance");
        }
        foreach (var value in graph.Values)
        {
            var valuePath = $"values[{value.LogicalId}]";
            RequireText(value.LogicalId, issues, "values.logicalId");
            RequireText(value.Name, issues, $"{valuePath}.name");
            RequireText(value.Kind, issues, $"{valuePath}.kind");
            RequireText(value.Value, issues, $"{valuePath}.value");
            RequireReference(value.TypeId, typeIds, issues, $"{valuePath}.typeId");
            foreach (var observationId in value.SupportObservationIds)
                RequireReference(observationId, observationIds, issues, $"{valuePath}.supportObservationIds");
            ValidateProvenance(value.Provenance, issues, $"{valuePath}.provenance");
        }
        foreach (var group in graph.AccessorGroups)
        {
            var groupPath = $"accessorGroups[{group.LogicalId}]";
            RequireText(group.LogicalId, issues, "accessorGroups.logicalId");
            RequireText(group.Name, issues, $"{groupPath}.name");
            RequireText(group.Kind, issues, $"{groupPath}.kind");
            RequireReference(group.TypeId, typeIds, issues, $"{groupPath}.typeId");
            if (group.EvidenceId is not null)
                RequireReference(group.EvidenceId, accessorEvidenceIds, issues, $"{groupPath}.evidenceId");
            foreach (var memberId in group.MemberIds)
            {
                RequireReference(memberId, memberIds, issues, $"{groupPath}.memberIds");
                if (membersById.TryGetValue(memberId, out var member)
                    && !string.Equals(member.AccessorGroupId, group.LogicalId, StringComparison.Ordinal))
                    issues.Add(new("accessor.member-mismatch", "Member and accessor group disagree.", $"{groupPath}.memberIds"));
            }
            ValidateProvenance(group.Provenance, issues, $"{groupPath}.provenance");
        }
        foreach (var observation in graph.SupportObservations)
        {
            var path = $"supportObservations[{observation.LogicalId}]";
            RequireText(observation.LogicalId, issues, "supportObservations.logicalId");
            RequireReference(observation.TargetId, allIds, issues, $"{path}.targetId");
            RequireText(observation.Product, issues, $"{path}.product");
            if (observation.Versions.Count == 0)
                issues.Add(new("support.versions", "A support observation must retain at least one source version.", $"{path}.versions"));
            foreach (var version in observation.Versions)
                RequireText(version, issues, $"{path}.versions");
            if (observation.LibraryIds.Count == 0)
                issues.Add(new("support.libraries", "A support observation must retain at least one source library.", $"{path}.libraryIds"));
            foreach (var libraryId in observation.LibraryIds)
                RequireReference(libraryId, libraryIds, issues, $"{path}.libraryIds");
            ValidateProvenance(observation.Provenance, issues, $"{path}.provenance");
        }
        foreach (var alias in graph.Aliases)
        {
            var path = $"aliases[{alias.LogicalId}]";
            RequireText(alias.LogicalId, issues, "aliases.logicalId");
            RequireText(alias.Alias, issues, $"{path}.alias");
            RequireReference(alias.TargetId, allIds, issues, $"{path}.targetId");
            RequireText(alias.Kind, issues, $"{path}.kind");
            ValidateProvenance(alias.Provenance, issues, $"{path}.provenance");
        }
        foreach (var unification in graph.Unifications)
        {
            var path = $"unifications[{unification.LogicalId}]";
            RequireText(unification.LogicalId, issues, "unifications.logicalId");
            RequireReference(unification.CanonicalId, allIds, issues, $"{path}.canonicalId");
            if (unification.EquivalentIds.Count == 0)
                issues.Add(new("unification.empty", "A unification must retain at least one equivalent identity.", path));
            foreach (var id in unification.EquivalentIds)
                RequireReference(id, allIds, issues, $"{path}.equivalentIds");
            RequireText(unification.Reason, issues, $"{path}.reason");
            ValidateProvenance(unification.Provenance, issues, $"{path}.provenance");
        }
        foreach (var evidence in graph.InvocationEvidence)
        {
            var path = $"invocationEvidence[{evidence.LogicalId}]";
            RequireText(evidence.LogicalId, issues, "invocationEvidence.logicalId");
            RequireReference(evidence.MemberId, memberIds, issues, $"{path}.memberId");
            RequireText(evidence.Operation, issues, $"{path}.operation");
            RequireText(evidence.DispatchName, issues, $"{path}.dispatchName");
            if (evidence.ArgumentCount < 0)
                issues.Add(new("invocation.argument-count", "ArgumentCount cannot be negative.", path));
            ValidateProvenance(evidence.Provenance, issues, $"{path}.provenance");
        }
        foreach (var evidence in graph.AccessorEvidence)
        {
            var path = $"accessorEvidence[{evidence.LogicalId}]";
            RequireText(evidence.LogicalId, issues, "accessorEvidence.logicalId");
            RequireReference(evidence.AccessorGroupId, groupIds, issues, $"{path}.accessorGroupId");
            RequireText(evidence.Kind, issues, $"{path}.kind");
            foreach (var memberId in evidence.MemberIds)
                RequireReference(memberId, memberIds, issues, $"{path}.memberIds");
            ValidateProvenance(evidence.Provenance, issues, $"{path}.provenance");
        }
        foreach (var absent in graph.AbsentFacts)
        {
            var path = $"absentFacts[{absent.LogicalId}]";
            RequireText(absent.LogicalId, issues, "absentFacts.logicalId");
            if (absent.TargetId is not null)
                RequireReference(absent.TargetId, allIds, issues, $"{path}.targetId");
            RequireText(absent.Kind, issues, $"{path}.kind");
            RequireText(absent.Reason, issues, $"{path}.reason");
            ValidateProvenance(absent.Provenance, issues, $"{path}.provenance");
            issues.Add(new("absent.record", "Absent source facts cannot be ignored during emission.", absent.LogicalId));
        }
        foreach (var unknown in graph.Unknowns)
        {
            RequireText(unknown.LogicalId, issues, "unknowns.logicalId");
            RequireText(unknown.Kind, issues, $"unknowns[{unknown.LogicalId}].kind");
            RequireText(unknown.Reason, issues, $"unknowns[{unknown.LogicalId}].reason");
            ValidateProvenance(unknown.Provenance, issues, $"unknowns[{unknown.LogicalId}].provenance");
            issues.Add(new("unknown.record", "Unknown facts cannot be emitted safely.", unknown.LogicalId));
        }
        foreach (var stale in graph.StaleRecords)
        {
            RequireText(stale.LogicalId, issues, "staleRecords.logicalId");
            RequireReference(stale.TargetId, allIds, issues, $"staleRecords[{stale.LogicalId}].targetId");
            ValidateSha(stale.ExpectedDigest, issues, $"staleRecords[{stale.LogicalId}].expectedDigest");
            ValidateSha(stale.ActualDigest, issues, $"staleRecords[{stale.LogicalId}].actualDigest");
            RequireText(stale.Reason, issues, $"staleRecords[{stale.LogicalId}].reason");
            ValidateProvenance(stale.Provenance, issues, $"staleRecords[{stale.LogicalId}].provenance");
            issues.Add(new("stale.record", "Stale facts must be refreshed before emission.", stale.LogicalId));
        }
        foreach (var ambiguity in graph.Ambiguities)
        {
            RequireText(ambiguity.LogicalId, issues, "ambiguities.logicalId");
            RequireText(ambiguity.Kind, issues, $"ambiguities[{ambiguity.LogicalId}].kind");
            RequireText(ambiguity.Reason, issues, $"ambiguities[{ambiguity.LogicalId}].reason");
            if (ambiguity.Candidates.Count == 1 || (ambiguity.Candidates.Count == 0 && !ambiguity.Kind.Contains("missing", StringComparison.OrdinalIgnoreCase)))
                issues.Add(new("ambiguity.candidates", "An ambiguity must retain multiple candidates; missing-fact records may have no candidates.", ambiguity.LogicalId));
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

    private static void ValidateIdentifiers(
        IEnumerable<DataIdentifierObservation> observations,
        IReadOnlySet<string> libraryIds,
        ICollection<ValidationIssue> issues,
        string path,
        bool requireGuid = false,
        bool requireInteger = false)
    {
        foreach (var observation in observations)
        {
            RequireText(observation.Value, issues, $"{path}.value");
            if (requireGuid && !Guid.TryParse(observation.Value, out _))
                issues.Add(new("identifier.guid", "Identifier observation is not a GUID.", $"{path}.value"));
            if (requireInteger && !int.TryParse(observation.Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out _))
                issues.Add(new("identifier.dispid", "DISPID observation is not an integer.", $"{path}.value"));
            if (observation.LibraryIds.Count == 0)
                issues.Add(new("identifier.libraries", "Identifier observation must identify at least one supporting library.", $"{path}.libraryIds"));
            foreach (var libraryId in observation.LibraryIds)
                RequireReference(libraryId, libraryIds, issues, $"{path}.libraryIds");
            ValidateProvenance(observation.Provenance, issues, $"{path}.provenance");
        }
    }

    private static void ValidateTypeReference(DataTypeReference reference, IReadOnlySet<string> typeIds, ICollection<ValidationIssue> issues, string path)
    {
        RequireText(reference.Name, issues, $"{path}.name");
        if (reference.IsExternal)
        {
            RequireText(reference.ProjectKey ?? "", issues, $"{path}.projectKey");
            RequireText(reference.LibraryKey ?? "", issues, $"{path}.libraryKey");
        }
        if (reference.TargetTypeId is not null)
            RequireReference(reference.TargetTypeId, typeIds, issues, $"{path}.targetTypeId");
        else if (reference.TypeKey is not null || reference.ProjectKey is not null)
            issues.Add(new("type-reference.unresolved", "A source type reference was not resolved to a Data v2 type.", path));
    }

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
