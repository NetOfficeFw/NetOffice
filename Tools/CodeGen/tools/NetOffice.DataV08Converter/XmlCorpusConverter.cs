using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using NetOffice.CodeGen.Data;

internal sealed record XmlProductStats(
    string Product,
    string Namespace,
    int SourceCategories,
    int Libraries,
    int Types,
    int Members,
    int Values,
    int SupportObservations,
    int SupportVersionFacts);

internal sealed record XmlConversionStats(
    int Files,
    int Projects,
    int Libraries,
    int Types,
    int Members,
    int Values,
    int Aliases,
    int SupportObservations,
    int SupportVersionFacts,
    int InvocationEvidence,
    int AccessorEvidence,
    int AbsentFacts,
    int Unknowns,
    int Ambiguities,
    int UnsupportedFiles,
    IReadOnlyList<XmlProductStats> Products);

internal sealed record XmlConversionResult(
    DataGraph Graph,
    XmlConversionStats Stats,
    IReadOnlyList<string> Diagnostics);

internal static class XmlCorpusConverter
{
    private static readonly Regex EncodedCharacter = new("_x([0-9a-fA-F]{4})_", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly HashSet<string> TypeElements = new(StringComparer.Ordinal)
    {
        "Alias", "CoClass", "Constant", "Enum", "Interface", "Module", "Record"
    };

    public static XmlConversionResult Convert(string inputDirectory, string sourcePath, string inputTreeHash, string? productFilter = null)
    {
        var root = Path.GetFullPath(inputDirectory);
        var diagnostics = new List<string>();
        var unknowns = new List<UnknownRecord>();
        var ambiguities = new List<AmbiguityRecord>();
        var absentFacts = new List<AbsentFactRecord>();
        var source = new DataSource
        {
            Kind = "exploratory-xml-unpinned",
            Path = NormalizePath(sourcePath),
            Sha256 = inputTreeHash
        };

        var libraryFile = Path.Combine(root, "Libraries.xml");
        if (!File.Exists(libraryFile))
            throw new InvalidDataException($"XML input is missing Libraries.xml: {libraryFile}");

        var libraryDocument = Load(libraryFile);
        var libraryDrafts = new List<LibraryDraft>();
        var librariesByKey = new Dictionary<string, DataLibrary>(StringComparer.Ordinal);
        foreach (var libraryNode in libraryDocument.Root?.Elements("Library") ?? Enumerable.Empty<XElement>())
        {
            var key = Attr(libraryNode, "Key");
            var guid = Decode(Attr(libraryNode, "GUID") ?? "");
            var name = Attr(libraryNode, "Name");
            if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(guid) || string.IsNullOrWhiteSpace(name))
            {
                AddUnknown(unknowns, source, "library-metadata", "Libraries.xml contains a library without Name, GUID, or Key.", libraryFile, libraryNode, root);
                continue;
            }
            if (!Guid.TryParse(guid, out _))
            {
                AddUnknown(unknowns, source, "library-guid", $"Library {name} has a non-GUID value: {guid}.", libraryFile, libraryNode, root);
                continue;
            }
            var logicalId = LogicalIds.Library(guid, key);
            if (librariesByKey.ContainsKey(key))
            {
                AddAmbiguity(ambiguities, source, "duplicate-library", key, [librariesByKey[key].LogicalId, logicalId], "Libraries.xml repeats a library key.", libraryFile, libraryNode, root);
                continue;
            }
            var library = new DataLibrary
            {
                LogicalId = logicalId,
                Name = name,
                Guid = guid,
                SourceKey = key,
                Version = Attr(libraryNode, "Version") ?? "",
                Major = Attr(libraryNode, "Major") ?? "",
                Minor = Attr(libraryNode, "Minor") ?? "",
                Description = NullIfEmpty(Attr(libraryNode, "Description")),
                Provenance = Provenance(source, libraryFile, libraryNode, root)
            };
            librariesByKey.Add(key, library);
            libraryDrafts.Add(new LibraryDraft(library, libraryNode));
        }

        var librariesByIdentity = librariesByKey.Values
            .GroupBy(static library => LibraryIdentity(library.Guid, library.Major, library.Minor), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(static group => group.Key, static group => group.Select(static item => item.LogicalId).OrderBy(static id => id, StringComparer.Ordinal).ToArray(), StringComparer.OrdinalIgnoreCase);
        foreach (var draft in libraryDrafts)
        {
            var dependencies = new List<DataLibraryDependency>();
            foreach (var dependencyNode in draft.Node.Elements("DependLib"))
            {
                var name = Attr(dependencyNode, "Name") ?? "";
                var guid = Decode(Attr(dependencyNode, "GUID") ?? "");
                var major = Attr(dependencyNode, "Major") ?? "";
                var minor = Attr(dependencyNode, "Minor") ?? "";
                var targetIds = librariesByIdentity.GetValueOrDefault(LibraryIdentity(guid, major, minor), Array.Empty<string>());
                if (string.IsNullOrWhiteSpace(name) || !Guid.TryParse(guid, out _) || string.IsNullOrWhiteSpace(major) || string.IsNullOrWhiteSpace(minor))
                    AddUnknown(unknowns, source, "library-dependency", $"Library {draft.Library.Name} has an incomplete dependency identity.", libraryFile, dependencyNode, root);
                else if (targetIds.Length == 0)
                    AddAmbiguity(ambiguities, source, "missing-library-dependency", LibraryIdentity(guid, major, minor), [], $"Dependency {name} {major}.{minor} is not present in Libraries.xml.", libraryFile, dependencyNode, root);
                dependencies.Add(new DataLibraryDependency
                {
                    Name = name,
                    Guid = guid,
                    Major = major,
                    Minor = minor,
                    Description = NullIfEmpty(Attr(dependencyNode, "Description")),
                    TargetLibraryIds = targetIds,
                    Provenance = Provenance(source, libraryFile, dependencyNode, root)
                });
            }
            librariesByKey[draft.Library.SourceKey] = draft.Library with { Dependencies = dependencies };
        }

        var projectDrafts = LoadProjects(root, source, librariesByKey, productFilter, unknowns, ambiguities);
        var projectsBySourceKey = projectDrafts.ToDictionary(static draft => draft.SourceKey, static draft => draft.Project.LogicalId, StringComparer.Ordinal);
        var projects = new List<DataProject>(projectDrafts.Count);
        foreach (var draft in projectDrafts)
        {
            var referenceIds = new List<string>();
            foreach (var referenceNode in draft.ReferenceNodes)
            {
                var key = Attr(referenceNode, "Key");
                if (string.IsNullOrWhiteSpace(key) || !projectsBySourceKey.TryGetValue(key, out var projectId))
                {
                    AddAmbiguity(ambiguities, source, "missing-project-reference", key ?? "missing", [], $"Project {draft.Project.Name} contains an unresolved project reference.", draft.RefProjectsFile, referenceNode, root);
                    continue;
                }
                referenceIds.Add(projectId);
            }
            projects.Add(draft.Project with { ReferenceProjectIds = referenceIds.Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray() });
        }
        var projectsById = projects.ToDictionary(static project => project.LogicalId, StringComparer.Ordinal);

        var typeById = new Dictionary<string, DataType>(StringComparer.Ordinal);
        var typeDrafts = new List<TypeDraft>();
        var typeIdsBySourceKey = new Dictionary<string, string>(StringComparer.Ordinal);
        var members = new List<DataMember>();
        var values = new List<DataValue>();
        var aliases = new List<AliasRecord>();
        var invocationEvidence = new List<InvocationEvidence>();
        var supportObservations = new Dictionary<string, SupportObservation>(StringComparer.Ordinal);
        var accessorGroups = new Dictionary<string, AccessorGroup>(StringComparer.Ordinal);
        var accessorMemberIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        foreach (var projectDraft in projectDrafts.Where(draft => productFilter is null || string.Equals(draft.Project.Name, productFilter, StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var includedFile in projectDraft.IncludedFiles.Where(static file => IsTypeSourceCategory(file.Category)))
            {
                if (!File.Exists(includedFile.Path))
                {
                    AddUnknown(unknowns, source, "project-source-missing", $"Project {projectDraft.Project.Name} includes missing source category {includedFile.RelativePath}.", projectDraft.ProjectFile, projectDraft.Node, root);
                    continue;
                }
                XDocument document;
                try
                {
                    document = Load(includedFile.Path);
                }
                catch (XmlException exception)
                {
                    AddUnknown(unknowns, source, "xml-parse", exception.Message, includedFile.Path, null, root);
                    diagnostics.Add($"Could not parse {includedFile.RelativePath}: {exception.Message}");
                    continue;
                }
                var rootElement = document.Root;
                if (rootElement is null)
                    continue;
                foreach (var typeNode in rootElement.Elements().Where(static node => TypeElements.Contains(node.Name.LocalName)))
                {
                    var typeName = Attr(typeNode, "Name");
                    var sourceKey = Attr(typeNode, "Key");
                    var ownerKey = FirstReferenceKey(typeNode.Element("RefLibraries"));
                    if (string.IsNullOrWhiteSpace(typeName) || string.IsNullOrWhiteSpace(sourceKey))
                    {
                        AddUnknown(unknowns, source, "type-identity", "A type is missing Name or Key; no identity was invented.", includedFile.Path, typeNode, root);
                        continue;
                    }
                    if (string.IsNullOrWhiteSpace(ownerKey) || !librariesByKey.TryGetValue(ownerKey, out var ownerLibrary))
                    {
                        AddAmbiguity(ambiguities, source, "missing-type-library", sourceKey, [], $"Type {typeName} has no resolvable owning library reference ({ownerKey ?? "missing"}).", includedFile.Path, typeNode, root);
                        continue;
                    }
                    var typeId = LogicalIds.Type(ownerLibrary.LogicalId, sourceKey);
                    if (typeById.ContainsKey(typeId) || typeIdsBySourceKey.ContainsKey(sourceKey))
                    {
                        var candidates = typeIdsBySourceKey.TryGetValue(sourceKey, out var existing) ? new[] { existing, typeId } : new[] { typeId };
                        AddAmbiguity(ambiguities, source, "duplicate-type", sourceKey, candidates, "The same type source identity occurs more than once.", includedFile.Path, typeNode, root);
                        continue;
                    }
                    var supportIds = AddSupportObservations(typeId, typeNode.Element("RefLibraries"), includedFile.Path, root, source, librariesByKey, supportObservations, ambiguities);
                    var type = new DataType
                    {
                        LogicalId = typeId,
                        LibraryId = ownerLibrary.LogicalId,
                        ProjectId = projectDraft.Project.LogicalId,
                        Namespace = projectDraft.Project.Namespace,
                        SourceCategory = includedFile.Category,
                        Name = typeName,
                        Kind = typeNode.Name.LocalName,
                        SourceKey = sourceKey,
                        AliasTarget = typeNode.Name.LocalName == "Alias" ? NullIfEmpty(Attr(typeNode, "Intrinsic")) : null,
                        DeclaredGuid = DecodeOptional(Attr(typeNode, "GUID")),
                        TypeLibType = ParseNullableInt(Attr(typeNode, "TypeLibType")),
                        IsEventInterface = ParseNullableBool(typeNode, "IsEventInterface"),
                        IsEarlyBind = ParseNullableBool(typeNode, "IsEarlyBind"),
                        IsHidden = ParseNullableBool(typeNode, "IsHidden"),
                        AutomaticQuit = ParseNullableBool(typeNode, "AutomaticQuit"),
                        IsApplicationObject = ParseNullableBool(typeNode, "IsAppObject"),
                        GuidObservations = ReadIdentifierObservations(typeNode.Element("DispIds"), includedFile.Path, root, source, librariesByKey, ambiguities),
                        SupportObservationIds = supportIds,
                        Provenance = Provenance(source, includedFile.Path, typeNode, root)
                    };
                    typeById.Add(typeId, type);
                    typeIdsBySourceKey.Add(sourceKey, typeId);
                    typeDrafts.Add(new TypeDraft(type, typeNode, includedFile.Path));
                    if (type.Kind == "Alias")
                    {
                        aliases.Add(new AliasRecord
                        {
                            LogicalId = LogicalIds.Alias($"{type.ProjectId}:{type.Name}"),
                            Alias = type.Name,
                            TargetId = type.LogicalId,
                            Kind = "type-alias",
                            Provenance = type.Provenance
                        });
                    }
                    ParseTypeMembers(type, typeNode, includedFile.Path, root, source, librariesByKey, supportObservations, members, values, invocationEvidence, accessorGroups, accessorMemberIds, unknowns, ambiguities);
                }
            }
        }

        foreach (var draft in typeDrafts)
        {
            var bases = ResolveTypeReferences(draft.Node.Element("Inherited"), draft, "base-type", root, source, librariesByKey, typeIdsBySourceKey, unknowns, ambiguities);
            var defaultInterfaces = ResolveTypeReferences(draft.Node.Element("DefaultInterfaces"), draft, "default-interface", root, source, librariesByKey, typeIdsBySourceKey, unknowns, ambiguities);
            var eventInterfaces = ResolveTypeReferences(draft.Node.Element("EventInterfaces"), draft, "event-interface", root, source, librariesByKey, typeIdsBySourceKey, unknowns, ambiguities);
            typeById[draft.Type.LogicalId] = draft.Type with
            {
                BaseTypeIds = bases.Ids,
                DefaultInterfaceIds = defaultInterfaces.Ids,
                EventInterfaceIds = eventInterfaces.Ids,
                ReferenceObservations = bases.Observations.Concat(defaultInterfaces.Observations).Concat(eventInterfaces.Observations).ToArray()
            };
        }
        var typeIdsByProjectAndName = typeById.Values
            .GroupBy(static type => (type.ProjectId, type.Name))
            .ToDictionary(static group => group.Key, static group => group.Select(static type => type.LogicalId).Single());
        for (var index = 0; index < members.Count; index++)
        {
            var member = members[index];
            var returnTypeReference = ResolveSignatureTypeReference(member.ReturnTypeReference, member.Provenance, source, projectsBySourceKey, typeIdsBySourceKey, typeIdsByProjectAndName, ambiguities);
            var parameters = member.Parameters.Select(parameter => parameter with
            {
                TypeReference = ResolveSignatureTypeReference(parameter.TypeReference, parameter.Provenance, source, projectsBySourceKey, typeIdsBySourceKey, typeIdsByProjectAndName, ambiguities)
            }).ToArray();
            members[index] = member with { ReturnTypeReference = returnTypeReference, Parameters = parameters };
        }


        var accessorEvidence = new List<AccessorEvidence>();
        foreach (var groupId in accessorGroups.Keys.OrderBy(static value => value, StringComparer.Ordinal).ToArray())
        {
            var group = accessorGroups[groupId];
            var memberIds = accessorMemberIds[groupId].OrderBy(static value => value, StringComparer.Ordinal).ToArray();
            var evidenceId = LogicalIds.Create("accessor-evidence", groupId);
            accessorEvidence.Add(new AccessorEvidence
            {
                LogicalId = evidenceId,
                AccessorGroupId = groupId,
                Kind = group.Kind,
                MemberIds = memberIds,
                Provenance = group.Provenance
            });
            accessorGroups[groupId] = group with { MemberIds = memberIds, EvidenceId = evidenceId };
        }

        var graph = new DataGraph
        {
            Source = source,
            Projects = projects,
            Libraries = librariesByKey.Values.ToArray(),
            Types = typeById.Values.ToArray(),
            Members = members,
            Values = values,
            AccessorGroups = accessorGroups.Values.ToArray(),
            SupportObservations = supportObservations.Values.ToArray(),
            Aliases = aliases,
            InvocationEvidence = invocationEvidence,
            AccessorEvidence = accessorEvidence,
            AbsentFacts = absentFacts,
            Unknowns = unknowns,
            Ambiguities = ambiguities
        };
        graph = graph with { Digest = CanonicalJson.ComputeDigest(graph) };

        var productStats = BuildProductStats(graph, projectsById);
        var stats = new XmlConversionStats(
            1 + (projectDrafts.Count * 3) + projectDrafts.Where(draft => productFilter is null || string.Equals(draft.Project.Name, productFilter, StringComparison.OrdinalIgnoreCase)).Sum(static draft => draft.IncludedFiles.Count(static file => IsTypeSourceCategory(file.Category))),
            graph.Projects.Count,
            graph.Libraries.Count,
            graph.Types.Count,
            graph.Members.Count,
            graph.Values.Count,
            graph.Aliases.Count,
            graph.SupportObservations.Count,
            graph.SupportObservations.Sum(static observation => observation.Versions.Count),
            graph.InvocationEvidence.Count,
            graph.AccessorEvidence.Count,
            graph.AbsentFacts.Count,
            graph.Unknowns.Count,
            graph.Ambiguities.Count,
            0,
            productStats);
        diagnostics.Add($"Converted {graph.Projects.Count} products from their explicit Project.xml namespaces and included source categories.");
        if (!string.IsNullOrWhiteSpace(productFilter))
            diagnostics.Add($"Exploratory product filter: {productFilter}");
        return new XmlConversionResult(graph, stats, diagnostics);
    }

    private static List<ProjectDraft> LoadProjects(
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        string? productFilter,
        ICollection<UnknownRecord> unknowns,
        ICollection<AmbiguityRecord> ambiguities)
    {
        var drafts = new List<ProjectDraft>();
        var seenKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var directory in Directory.EnumerateDirectories(root).OrderBy(static path => path, StringComparer.Ordinal))
        {
            var projectFile = Path.Combine(directory, "Project.xml");
            if (!File.Exists(projectFile))
                continue;
            var document = Load(projectFile);
            var node = document.Root;
            if (node is null || node.Name.LocalName != "Project")
            {
                AddUnknown(unknowns, source, "project-metadata", "Project.xml has no Project root.", projectFile, node, root);
                continue;
            }
            var name = Attr(node, "Name");
            var projectNamespace = Attr(node, "Namespace");
            var sourceKey = Attr(node, "Key");
            var version = Attr(node, "Version");
            var fileVersion = Attr(node, "FileVersion");
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(projectNamespace) || string.IsNullOrWhiteSpace(sourceKey) || string.IsNullOrWhiteSpace(version) || string.IsNullOrWhiteSpace(fileVersion))
            {
                AddUnknown(unknowns, source, "project-metadata", "Project.xml is missing Name, Namespace, Key, Version, or FileVersion.", projectFile, node, root);
                continue;
            }
            if (!seenKeys.Add(sourceKey))
            {
                AddAmbiguity(ambiguities, source, "duplicate-project", sourceKey, [], $"Project key {sourceKey} occurs more than once.", projectFile, node, root);
                continue;
            }
            var includedFiles = node.Elements()
                .Where(static element => element.Name.LocalName == "include")
                .Select(element => Attr(element, "href"))
                .Where(static href => !string.IsNullOrWhiteSpace(href))
                .Select(href => new IncludedFile(
                    Path.Combine(directory, href!),
                    NormalizeRelative(root, Path.Combine(directory, href!)),
                    Path.GetFileNameWithoutExtension(href!)))
                .ToArray();
            if (includedFiles.Length == 0)
                AddUnknown(unknowns, source, "project-source-categories", $"Project {name} contains no XML source includes.", projectFile, node, root);

            var refLibrariesFile = Path.Combine(directory, "RefLibraries.xml");
            var libraryIds = new List<string>();
            if (!File.Exists(refLibrariesFile))
                AddUnknown(unknowns, source, "project-libraries", $"Project {name} is missing RefLibraries.xml.", projectFile, node, root);
            else
            {
                foreach (var referenceNode in Load(refLibrariesFile).Root?.Elements("Ref") ?? Enumerable.Empty<XElement>())
                {
                    var key = Attr(referenceNode, "Key");
                    if (string.IsNullOrWhiteSpace(key) || !librariesByKey.TryGetValue(key, out var library))
                    {
                        AddAmbiguity(ambiguities, source, "missing-project-library", key ?? "missing", [], $"Project {name} contains an unresolved library reference.", refLibrariesFile, referenceNode, root);
                        continue;
                    }
                    libraryIds.Add(library.LogicalId);
                }
            }

            var refProjectsFile = Path.Combine(directory, "RefProjects.xml");
            IReadOnlyList<XElement> referenceNodes;
            if (!File.Exists(refProjectsFile))
            {
                AddUnknown(unknowns, source, "project-references", $"Project {name} is missing RefProjects.xml.", projectFile, node, root);
                referenceNodes = Array.Empty<XElement>();
            }
            else
                referenceNodes = (Load(refProjectsFile).Root?.Elements("RefProject") ?? Enumerable.Empty<XElement>()).ToArray();

            var project = new DataProject
            {
                LogicalId = LogicalIds.Project(sourceKey),
                Name = name,
                Namespace = projectNamespace,
                SourceKey = sourceKey,
                Version = version,
                FileVersion = fileVersion,
                Ignore = IsTrue(node, "Ignore"),
                SourceCategories = includedFiles.Select(static file => file.Category).Distinct(StringComparer.Ordinal).OrderBy(static category => category, StringComparer.Ordinal).ToArray(),
                LibraryIds = libraryIds.Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                Provenance = Provenance(source, projectFile, node, root)
            };
            drafts.Add(new ProjectDraft(project, sourceKey, node, projectFile, refProjectsFile, includedFiles, referenceNodes));
        }
        if (drafts.Count == 0)
            throw new InvalidDataException($"XML input contains no product Project.xml files under {root}.");
        if (!string.IsNullOrWhiteSpace(productFilter) && !drafts.Any(draft => string.Equals(draft.Project.Name, productFilter, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException($"Product filter did not match a Project.xml under {root}: {productFilter}");
        return drafts;
    }

    private static void ParseTypeMembers(
        DataType type,
        XElement typeNode,
        string file,
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        IDictionary<string, SupportObservation> supportObservations,
        ICollection<DataMember> members,
        ICollection<DataValue> values,
        ICollection<InvocationEvidence> invocationEvidence,
        IDictionary<string, AccessorGroup> accessorGroups,
        IDictionary<string, List<string>> accessorMemberIds,
        ICollection<UnknownRecord> unknowns,
        ICollection<AmbiguityRecord> ambiguities)
    {
        var valueEntries = typeNode.Element("Members")?.Elements("Member") ?? Enumerable.Empty<XElement>();
        if (type.Kind == "Record")
        {
            var fieldOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var fieldNode in valueEntries)
            {
                var name = Attr(fieldNode, "Name");
                var fieldType = Attr(fieldNode, "Type");
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(fieldType))
                {
                    AddUnknown(unknowns, source, "record-field", "Record field is missing Name or Type; no field signature was invented.", file, fieldNode, root);
                    continue;
                }
                fieldOrdinals.TryGetValue(name, out var ordinal);
                fieldOrdinals[name] = ordinal + 1;
                var key = ordinal == 0 ? $"field:{name}" : $"field:{name}#{ordinal + 1}";
                var provenance = Provenance(source, file, fieldNode, root);
                var memberId = LogicalIds.Member(type.LogicalId, key);
                var supportIds = AddSupportObservations(memberId, fieldNode.Element("RefLibraries"), file, root, source, librariesByKey, supportObservations, ambiguities);
                members.Add(new DataMember
                {
                    LogicalId = memberId,
                    TypeId = type.LogicalId,
                    Name = name,
                    Kind = "field",
                    SourceKey = key,
                    ReturnType = fieldType,
                    ReturnTypeReference = ReadTypeReference(fieldNode),
                    IsComProxy = IsTrue(fieldNode, "IsComProxy"),
                    SupportObservationIds = supportIds,
                    Provenance = provenance
                });
            }
            return;
        }

        var valueOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var valueNode in valueEntries)
        {
            var name = Attr(valueNode, "Name");
            var value = Attr(valueNode, "Value");
            if (string.IsNullOrWhiteSpace(name))
            {
                AddUnknown(unknowns, source, "value-identity", "A value is missing Name; no identity was invented.", file, valueNode, root);
                continue;
            }
            if (value is null)
            {
                AddUnknown(unknowns, source, "value-fact", $"Value {name} has no Value attribute.", file, valueNode, root);
                continue;
            }
            valueOrdinals.TryGetValue(name, out var ordinal);
            valueOrdinals[name] = ordinal + 1;
            var key = ordinal == 0 ? $"value:{name}" : $"value:{name}#{ordinal + 1}";
            var valueId = LogicalIds.Member(type.LogicalId, key);
            var provenance = Provenance(source, file, valueNode, root);
            var supportIds = AddSupportObservations(valueId, valueNode.Element("RefLibraries"), file, root, source, librariesByKey, supportObservations, ambiguities);
            values.Add(new DataValue
            {
                LogicalId = valueId,
                TypeId = type.LogicalId,
                Name = name,
                Kind = type.Kind,
                Value = value,
                ValueType = NullIfEmpty(Attr(valueNode, "Type")),
                SupportObservationIds = supportIds,
                Provenance = provenance
            });
        }

        var callableNodes = (typeNode.Element("Methods")?.Elements("Method") ?? Enumerable.Empty<XElement>())
            .Concat(typeNode.Element("Properties")?.Elements("Property") ?? Enumerable.Empty<XElement>())
            .Concat(typeNode.Element("Events")?.Elements("Event") ?? Enumerable.Empty<XElement>())
            .ToArray();
        var memberOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var memberNode in callableNodes)
        {
            var name = Attr(memberNode, "Name");
            var sourceKey = Attr(memberNode, "Key");
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(sourceKey))
            {
                AddUnknown(unknowns, source, "member-identity", "A callable member is missing Name or Key; no identity was invented.", file, memberNode, root);
                continue;
            }
            memberOrdinals.TryGetValue(sourceKey, out var ordinal);
            memberOrdinals[sourceKey] = ordinal + 1;
            var identityKey = ordinal == 0 ? sourceKey : $"{sourceKey}#{ordinal + 1}";
            var provenance = Provenance(source, file, memberNode, root);
            var parameterResult = ReadParameters(memberNode.Element("Parameters"), file, root, source, librariesByKey, unknowns, ambiguities);
            var dispIdObservations = ReadIdentifierObservations(memberNode.Element("DispIds"), file, root, source, librariesByKey, ambiguities);
            var dispId = ReadFirstDispId(dispIdObservations, file, root, source, memberNode, unknowns);
            var accessorKind = ReadAccessorKind(memberNode);
            var accessorGroupId = memberNode.Name.LocalName is "Property" or "Event"
                ? LogicalIds.AccessorGroup(type.LogicalId, name)
                : null;
            var memberId = LogicalIds.Member(type.LogicalId, identityKey);
            var supportIds = AddSupportObservations(memberId, memberNode.Element("RefLibraries"), file, root, source, librariesByKey, supportObservations, ambiguities);
            var evidenceId = LogicalIds.Create("invocation", memberId);
            invocationEvidence.Add(new InvocationEvidence
            {
                LogicalId = evidenceId,
                MemberId = memberId,
                Operation = memberNode.Name.LocalName.ToLowerInvariant(),
                DispatchName = name,
                DispId = dispId,
                ArgumentCount = parameterResult.Parameters.Count,
                ResultType = parameterResult.ReturnType,
                RequiresProxy = parameterResult.RequiresProxy,
                Provenance = provenance
            });
            members.Add(new DataMember
            {
                LogicalId = memberId,
                TypeId = type.LogicalId,
                Name = name,
                Kind = memberNode.Name.LocalName.ToLowerInvariant(),
                SourceKey = sourceKey,
                DispId = dispId,
                DispIdObservations = dispIdObservations,
                ReturnType = parameterResult.ReturnType,
                ReturnTypeReference = parameterResult.ReturnTypeReference,
                IsHidden = IsTrue(memberNode, "Hidden"),
                AnalyzeReturn = IsTrue(memberNode, "AnalyzeReturn"),
                ParameterTypes = parameterResult.Parameters.Select(static parameter => parameter.Type).ToArray(),
                Parameters = parameterResult.Parameters,
                SupportObservationIds = supportIds,
                SignatureLibraryIds = parameterResult.SignatureLibraryIds,
                AccessorGroupId = accessorGroupId,
                AccessorKind = accessorKind,
                InvocationEvidenceId = evidenceId,
                Provenance = provenance
            });
            if (accessorGroupId is not null)
            {
                if (!accessorGroups.TryGetValue(accessorGroupId, out var group))
                    accessorGroups[accessorGroupId] = group = new AccessorGroup
                    {
                        LogicalId = accessorGroupId,
                        TypeId = type.LogicalId,
                        Name = name,
                        Kind = memberNode.Name.LocalName.ToLowerInvariant(),
                        Provenance = provenance
                    };
                if (!accessorMemberIds.TryGetValue(accessorGroupId, out var ids))
                    accessorMemberIds[accessorGroupId] = ids = new List<string>();
                ids.Add(memberId);
            }
            if (memberNode.Name.LocalName == "Property" && accessorKind is null)
                AddUnknown(unknowns, source, "accessor-kind", $"Property {name} has no recognized InvokeKind; access direction was not invented.", file, memberNode, root);
        }
    }

    private static IReadOnlyList<string> AddSupportObservations(
        string targetId,
        XElement? references,
        string file,
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        IDictionary<string, SupportObservation> observations,
        ICollection<AmbiguityRecord> ambiguities)
    {
        if (references is null)
            return Array.Empty<string>();
        var ids = new List<string>();
        foreach (var reference in references.Elements("Ref"))
        {
            var key = Attr(reference, "Key");
            if (string.IsNullOrWhiteSpace(key) || !librariesByKey.TryGetValue(key, out var library))
            {
                AddAmbiguity(ambiguities, source, "missing-support-library", key ?? "missing", [], $"Support observation for {targetId} references an unknown library.", file, reference, root);
                continue;
            }
            var observationId = LogicalIds.Create("support-observation", targetId, library.Name);
            if (observations.TryGetValue(observationId, out var existing))
            {
                observations[observationId] = existing with
                {
                    Versions = existing.Versions.Append(library.Version).Distinct(StringComparer.Ordinal).OrderBy(static version => version, StringComparer.Ordinal).ToArray(),
                    LibraryIds = existing.LibraryIds.Append(library.LogicalId).Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray()
                };
            }
            else
            {
                observations.Add(observationId, new SupportObservation
                {
                    LogicalId = observationId,
                    TargetId = targetId,
                    Product = library.Name,
                    Versions = new[] { library.Version },
                    LibraryIds = new[] { library.LogicalId },
                    Present = true,
                    Provenance = Provenance(source, file, reference, root)
                });
            }
            ids.Add(observationId);
        }
        return ids.Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<string> ResolveLibraryIds(
        XElement? references,
        string file,
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        ICollection<AmbiguityRecord> ambiguities,
        string kind)
    {
        if (references is null)
            return Array.Empty<string>();
        var ids = new List<string>();
        foreach (var reference in references.Elements("Ref"))
        {
            var key = Attr(reference, "Key");
            if (string.IsNullOrWhiteSpace(key) || !librariesByKey.TryGetValue(key, out var library))
            {
                AddAmbiguity(ambiguities, source, $"missing-{kind}-library", key ?? "missing", [], $"{kind} references an unknown library.", file, reference, root);
                continue;
            }
            ids.Add(library.LogicalId);
        }
        return ids.Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray();
    }

    private static ParameterResult ReadParameters(
        XElement? parametersNode,
        string file,
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        ICollection<UnknownRecord> unknowns,
        ICollection<AmbiguityRecord> ambiguities)
    {
        if (parametersNode is null)
        {
            AddUnknown(unknowns, source, "signature", "Callable member has no Parameters element; signature was not invented.", file, null, root);
            return new ParameterResult(null, null, Array.Empty<DataParameter>(), Array.Empty<string>(), false);
        }
        var returnNode = parametersNode.Element("ReturnValue");
        var returnType = Attr(returnNode, "Type");
        if (returnNode is null || string.IsNullOrWhiteSpace(returnType))
            AddUnknown(unknowns, source, "return-type", "Callable member has no return Type attribute.", file, returnNode, root);
        var returnReference = returnNode is null ? null : ReadTypeReference(returnNode);
        var result = new List<DataParameter>();
        var requiresProxy = returnReference?.IsComProxy == true;
        var index = 0;
        foreach (var parameterNode in parametersNode.Elements("Parameter"))
        {
            var name = Attr(parameterNode, "Name");
            var type = Attr(parameterNode, "Type");
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(type))
            {
                AddUnknown(unknowns, source, "parameter-fact", "Parameter is missing Name or Type; no signature fact was invented.", file, parameterNode, root);
                continue;
            }
            var isOut = IsTrue(parameterNode, "IsOut");
            var isRef = IsTrue(parameterNode, "IsRef");
            var optional = IsTrue(parameterNode, "IsOptional");
            var hasDefault = IsTrue(parameterNode, "HasDefaultValue");
            var reference = ReadTypeReference(parameterNode);
            requiresProxy |= reference.IsComProxy;
            result.Add(new DataParameter
            {
                Name = name,
                Type = type,
                RefKind = isOut ? "out" : isRef ? "ref" : "value",
                IsOptional = optional,
                HasDefaultValue = hasDefault,
                DefaultValue = hasDefault ? Attr(parameterNode, "DefaultValue") : null,
                TypeReference = reference,
                ParamFlags = NullIfEmpty(Attr(parameterNode, "ParamFlags")),
                Provenance = Provenance(source, file, parameterNode, root, $"parameters/{index}")
            });
            index++;
        }
        var signatureLibraryIds = ResolveLibraryIds(parametersNode.Element("RefLibraries"), file, root, source, librariesByKey, ambiguities, "signature");
        return new ParameterResult(returnType, returnReference, result, signatureLibraryIds, requiresProxy);
    }

    private static DataTypeReference ReadTypeReference(XElement node)
        => new()
        {
            Name = Attr(node, "Type") ?? "",
            TypeKind = NullIfEmpty(Attr(node, "TypeKind")),
            VarType = NullIfEmpty(Attr(node, "VarType")),
            MarshalAs = NullIfEmpty(Attr(node, "MarshalAs")),
            TypeKey = NullIfEmpty(Attr(node, "TypeKey")),
            ProjectKey = NullIfEmpty(Attr(node, "ProjectKey")),
            LibraryKey = NullIfEmpty(Attr(node, "LibraryKey")),
            IsComProxy = IsTrue(node, "IsComProxy"),
            IsExternal = IsTrue(node, "IsExternal"),
            IsEnum = IsTrue(node, "IsEnum"),
            IsArray = IsTrue(node, "IsArray"),
            IsNative = IsTrue(node, "IsNative")
        };

    private static IReadOnlyList<DataIdentifierObservation> ReadIdentifierObservations(
        XElement? identifiersNode,
        string file,
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        ICollection<AmbiguityRecord> ambiguities)
    {
        if (identifiersNode is null)
            return Array.Empty<DataIdentifierObservation>();
        var observations = new List<DataIdentifierObservation>();
        foreach (var identifierNode in identifiersNode.Elements("DispId"))
        {
            var value = Decode(Attr(identifierNode, "Id") ?? "");
            if (string.IsNullOrWhiteSpace(value))
                continue;
            var libraryIds = new List<string>();
            foreach (var reference in identifierNode.Element("RefLibraries")?.Elements("Ref") ?? Enumerable.Empty<XElement>())
            {
                var key = Attr(reference, "Key");
                if (string.IsNullOrWhiteSpace(key) || !librariesByKey.TryGetValue(key, out var library))
                {
                    AddAmbiguity(ambiguities, source, "missing-identifier-library", key ?? "missing", [], $"Identifier {value} references an unknown library.", file, reference, root);
                    continue;
                }
                libraryIds.Add(library.LogicalId);
            }
            observations.Add(new DataIdentifierObservation
            {
                Value = value,
                LibraryIds = libraryIds.Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                Provenance = Provenance(source, file, identifierNode, root)
            });
        }
        return observations
            .GroupBy(static observation => observation.Value + "\0" + string.Join("\0", observation.LibraryIds), StringComparer.Ordinal)
            .Select(static group => group.First())
            .ToArray();
    }

    private static int? ReadFirstDispId(
        IReadOnlyList<DataIdentifierObservation> observations,
        string file,
        string root,
        DataSource source,
        XElement memberNode,
        ICollection<UnknownRecord> unknowns)
    {
        if (observations.Count == 0)
        {
            AddUnknown(unknowns, source, "dispid", "Callable member has no DISPID observation.", file, memberNode, root);
            return null;
        }
        if (int.TryParse(observations[0].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            return value;
        AddUnknown(unknowns, source, "dispid-format", $"DISPID value '{observations[0].Value}' is not an integer after XML escape decoding.", file, memberNode, root);
        return null;
    }

    private static ResolvedTypeReferences ResolveTypeReferences(
        XElement? references,
        TypeDraft draft,
        string kind,
        string root,
        DataSource source,
        IReadOnlyDictionary<string, DataLibrary> librariesByKey,
        IReadOnlyDictionary<string, string> typeIdsBySourceKey,
        ICollection<UnknownRecord> unknowns,
        ICollection<AmbiguityRecord> ambiguities)
    {
        if (references is null)
            return new ResolvedTypeReferences(Array.Empty<string>(), Array.Empty<DataTypeLinkObservation>());
        var ids = new List<string>();
        var observations = new List<DataTypeLinkObservation>();
        foreach (var reference in references.Elements("Ref"))
        {
            var key = Attr(reference, "Key");
            if (string.IsNullOrWhiteSpace(key))
            {
                AddUnknown(unknowns, source, $"{kind}-key", $"A {kind} reference has no Key.", draft.File, reference, root);
                continue;
            }
            if (!typeIdsBySourceKey.TryGetValue(key, out var targetId))
            {
                AddAmbiguity(ambiguities, source, $"missing-{kind}", key, [], $"Referenced {kind} {key} was not found in the XML corpus.", draft.File, reference, root);
                continue;
            }
            var libraryIds = ResolveLibraryIds(reference.Element("RefLibraries"), draft.File, root, source, librariesByKey, ambiguities, kind);
            ids.Add(targetId);
            observations.Add(new DataTypeLinkObservation
            {
                Kind = kind,
                TargetTypeId = targetId,
                LibraryIds = libraryIds,
                Provenance = Provenance(source, draft.File, reference, root)
            });
        }
        return new ResolvedTypeReferences(
            ids.Distinct(StringComparer.Ordinal).OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
            observations);
    }

    private static DataTypeReference? ResolveSignatureTypeReference(
        DataTypeReference? reference,
        Provenance provenance,
        DataSource source,
        IReadOnlyDictionary<string, string> projectIdsBySourceKey,
        IReadOnlyDictionary<string, string> typeIdsBySourceKey,
        IReadOnlyDictionary<(string ProjectId, string Name), string> typeIdsByProjectAndName,
        ICollection<AmbiguityRecord> ambiguities)
    {
        if (reference is null)
            return null;
        string? targetTypeId = null;
        if (reference.TypeKey is not null)
            typeIdsBySourceKey.TryGetValue(reference.TypeKey, out targetTypeId);
        else if (reference.ProjectKey is not null
            && projectIdsBySourceKey.TryGetValue(reference.ProjectKey, out var projectId))
            typeIdsByProjectAndName.TryGetValue((projectId, reference.Name), out targetTypeId);
        if (targetTypeId is not null)
            return reference with { TargetTypeId = targetTypeId };
        if (reference.TypeKey is not null || reference.ProjectKey is not null)
        {
            var identity = reference.TypeKey ?? $"{reference.ProjectKey}:{reference.Name}";
            ambiguities.Add(new AmbiguityRecord
            {
                LogicalId = LogicalIds.Ambiguity("missing-signature-type", $"{identity}:{provenance.Location}"),
                Kind = "missing-signature-type",
                Reason = $"Signature type reference {identity} could not be resolved without inventing a type.",
                Candidates = Array.Empty<string>(),
                Provenance = provenance with { SourcePath = source.Path, SourceSha256 = source.Sha256 }
            });
        }
        return reference;
    }

    private static IReadOnlyList<XmlProductStats> BuildProductStats(DataGraph graph, IReadOnlyDictionary<string, DataProject> projectsById)
    {
        var typesByProject = graph.Types.GroupBy(static type => type.ProjectId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.ToArray(), StringComparer.Ordinal);
        var membersByType = graph.Members.GroupBy(static member => member.TypeId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        var valuesByType = graph.Values.GroupBy(static value => value.TypeId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        var observationsByTarget = graph.SupportObservations.GroupBy(static observation => observation.TargetId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        var supportVersionsByTarget = graph.SupportObservations.GroupBy(static observation => observation.TargetId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Sum(static observation => observation.Versions.Count), StringComparer.Ordinal);
        var result = new List<XmlProductStats>();
        foreach (var project in projectsById.Values.OrderBy(static project => project.Name, StringComparer.Ordinal))
        {
            var types = typesByProject.GetValueOrDefault(project.LogicalId, Array.Empty<DataType>());
            var typeIds = types.Select(static type => type.LogicalId).ToHashSet(StringComparer.Ordinal);
            var productMembers = graph.Members.Where(member => typeIds.Contains(member.TypeId));
            var productValues = graph.Values.Where(value => typeIds.Contains(value.TypeId));
            var memberCount = types.Sum(type => membersByType.GetValueOrDefault(type.LogicalId));
            var valueCount = types.Sum(type => valuesByType.GetValueOrDefault(type.LogicalId));
            var supportCount = types.Sum(type => observationsByTarget.GetValueOrDefault(type.LogicalId));
            supportCount += productMembers.Sum(member => observationsByTarget.GetValueOrDefault(member.LogicalId));
            supportCount += productValues.Sum(value => observationsByTarget.GetValueOrDefault(value.LogicalId));
            var supportVersionCount = types.Sum(type => supportVersionsByTarget.GetValueOrDefault(type.LogicalId));
            supportVersionCount += productMembers.Sum(member => supportVersionsByTarget.GetValueOrDefault(member.LogicalId));
            supportVersionCount += productValues.Sum(value => supportVersionsByTarget.GetValueOrDefault(value.LogicalId));
            result.Add(new XmlProductStats(project.Name, project.Namespace, project.SourceCategories.Count, project.LibraryIds.Count, types.Length, memberCount, valueCount, supportCount, supportVersionCount));
        }
        return result;
    }

    private static string? ReadAccessorKind(XElement node)
    {
        var invokeKind = Attr(node, "InvokeKind")?.ToUpperInvariant();
        return invokeKind switch
        {
            "INVOKE_PROPERTYGET" => "get",
            "INVOKE_PROPERTYPUT" => "put",
            "INVOKE_PROPERTYPUTREF" => "putref",
            _ => null
        };
    }

    private static string? FirstReferenceKey(XElement? refs)
        => Attr(refs?.Elements("Ref").FirstOrDefault(), "Key");

    private static XDocument Load(string path)
    {
        using var stream = File.OpenRead(path);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XDocument.Load(reader, LoadOptions.SetLineInfo);
    }

    private static string? Attr(XElement? element, string name)
        => element?.Attribute(name)?.Value;

    private static bool IsTrue(XElement? element, string name)
        => string.Equals(Attr(element, name), "true", StringComparison.OrdinalIgnoreCase);

    private static bool? ParseNullableBool(XElement element, string name)
        => element.Attribute(name) is null ? null : IsTrue(element, name);

    private static int? ParseNullableInt(string? value)
        => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;

    private static string Decode(string value)
        => EncodedCharacter.Replace(value, static match => ((char)int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());

    private static string? DecodeOptional(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : Decode(value);

    private static string? NullIfEmpty(string? value)
        => string.IsNullOrEmpty(value) ? null : value;

    private static string LibraryIdentity(string guid, string major, string minor)
        => $"{guid.ToUpperInvariant()}\0{major}\0{minor}";

    private static bool IsTypeSourceCategory(string category)
        => category is "Constants" or "Enums" or "Modules" or "TypeDefs" or "Records" or "CoClasses" or "DispatchInterfaces" or "Interfaces";

    private static string NormalizePath(string path)
        => path.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');

    private static string NormalizeRelative(string root, string path)
        => NormalizePath(Path.GetRelativePath(root, path));

    private static Provenance Provenance(DataSource source, string file, XElement? node, string root, string? suffix = null)
    {
        var location = NormalizeRelative(root, file);
        if (node is IXmlLineInfo lineInfo && lineInfo.HasLineInfo())
            location += ":" + lineInfo.LineNumber.ToString(CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(suffix))
            location += "/" + suffix;
        return new Provenance { SourcePath = source.Path, SourceSha256 = source.Sha256, Location = location };
    }

    private static void AddUnknown(ICollection<UnknownRecord> records, DataSource source, string kind, string reason, string file, XElement? node, string? root = null)
    {
        var location = root is null ? NormalizePath(file) : Provenance(source, file, node, root).Location ?? NormalizePath(file);
        records.Add(new UnknownRecord
        {
            LogicalId = LogicalIds.Create("unknown", kind, location, reason),
            Kind = kind,
            Reason = reason,
            Provenance = new Provenance { SourcePath = source.Path, SourceSha256 = source.Sha256, Location = location }
        });
    }

    private static void AddAmbiguity(ICollection<AmbiguityRecord> records, DataSource source, string kind, string key, IReadOnlyList<string> candidates, string reason, string file, XElement? node, string? root = null)
    {
        var location = root is null ? NormalizePath(file) : Provenance(source, file, node, root).Location;
        records.Add(new AmbiguityRecord
        {
            LogicalId = LogicalIds.Ambiguity(kind, $"{key}:{location}"),
            Kind = kind,
            Reason = reason,
            Candidates = candidates,
            Provenance = new Provenance { SourcePath = source.Path, SourceSha256 = source.Sha256, Location = location }
        });
    }

    private sealed record LibraryDraft(DataLibrary Library, XElement Node);
    private sealed record IncludedFile(string Path, string RelativePath, string Category);
    private sealed record ProjectDraft(
        DataProject Project,
        string SourceKey,
        XElement Node,
        string ProjectFile,
        string RefProjectsFile,
        IReadOnlyList<IncludedFile> IncludedFiles,
        IReadOnlyList<XElement> ReferenceNodes);
    private sealed record TypeDraft(DataType Type, XElement Node, string File);
    private sealed record ParameterResult(
        string? ReturnType,
        DataTypeReference? ReturnTypeReference,
        IReadOnlyList<DataParameter> Parameters,
        IReadOnlyList<string> SignatureLibraryIds,
        bool RequiresProxy);
    private sealed record ResolvedTypeReferences(IReadOnlyList<string> Ids, IReadOnlyList<DataTypeLinkObservation> Observations);
}
