using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using NetOffice.CodeGen.Data;

namespace NetOffice.CodeGen.Projection;

/// <summary>Pure, deterministic projection from the complete Data v2 graph, overlaid by the closed wrapper parity contract.</summary>
public static class ProjectionEngine
{
    private static readonly Regex SupportRegex = new(@"SupportByVersion\s*\(\s*\""(?<product>[^\""\r\n]+)\""\s*,(?<versions>[^)]*)\)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex AttributeRegex = new("[A-Za-z_][A-Za-z0-9_]*", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex DelegateRegex = new(@"\bdelegate\s+(?<return>[^\s]+)\s+[A-Za-z_][A-Za-z0-9_]*\s*\((?<parameters>.*)\)\s*;?$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex IndexPropertyRegex = new(@"HasIndexProperty\s*\([^,]+,\s*\""(?<name>[^\""]+)\""\s*\)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex SinkArgumentRegex = new(@"SinkArgument\s*\(\s*\""(?<name>[^\""]+)\""\s*,\s*(?<kind>SinkArgumentType\.[A-Za-z0-9_]+|typeof\((?<type>[^)]+)\))(?:\s*,\s*typeof\((?<enum>[^)]+)\))?\s*\)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex RedirectRegex = new(@"Redirect\s*\(\s*\""(?<target>[^\""]+)\""\s*\)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex ContractInvocationRegex = new(@"Execute[A-Za-z0-9]*?(?:MethodGet|PropertyGet|PropertySet)(?:<[^>]+>)?\s*\(\s*(?<target>[A-Za-z_][A-Za-z0-9_.]*)\s*,\s*\""(?<dispatch>[^\""]+)\""\s*(?:,\s*(?<arguments>.*))?\)\s*;?", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex EventBackingFieldRegex = new(@"(?<field>_[A-Za-z_][A-Za-z0-9_]*)\s*\+=", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex SinkLocalDeclarationRegex = new(@"^\s*(?<type>[A-Za-z_][A-Za-z0-9_:.<>\[\]]*)\s+(?<local>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?<expression>.+);\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex InvokerDispatchRegex = new(@"Invoker\.(?:MethodReturn|PropertyGet|PropertySet)\s*\(\s*(?<target>[A-Za-z_][A-Za-z0-9_.]*)\s*,\s*\""(?<dispatch>[^\""]+)\""", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex ValidateParamsRegex = new(@"Invoker\.ValidateParamsArray\s*\((?<arguments>.*)\)\s*;?", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex EventValidationRegex = new(@"Validate\s*\(\s*\""(?<key>[^\""]+)\""\s*\)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex ReleaseParamsRegex = new(@"(?<![""A-Za-z0-9_])Invoker\.ReleaseParamsArray\s*\((?<arguments>[^)]*)\)\s*;?", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly HashSet<string> ScalarTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "bool", "boolean", "byte", "sbyte", "short", "int16", "ushort", "uint16", "int", "int32", "uint", "uint32", "long", "int64", "ulong", "uint64", "single", "float", "double", "decimal", "char", "string", "datetime", "intptr", "uintptr"
    };

    private static readonly IReadOnlyDictionary<string, DataType[]> EmptyTypes = new Dictionary<string, DataType[]>(StringComparer.Ordinal);
    private static readonly IReadOnlySet<string> EmptyStringSet = new HashSet<string>(StringComparer.Ordinal);
    private static readonly IReadOnlySet<string> RetainedExactContractMemberNames = new HashSet<string>(StringComparer.Ordinal)
    {
        "InstanceType", "LateBindingApiWrapperType",
        "CreateEventBridge", "EventBridgeInitialized", "HasEventRecipients", "GetEventRecipients",
        "GetCountOfEventRecipients", "RaiseCustomEvent", "DisposeEventBridge",
        "AssemblyName", "AssemblyNamespace", "ComponentGuid", "AssemblyAttribute", "Assembly",
        "Dependencies", "Contains", "GetComObjectEnumerator", "FetchVariantComObjectEnumerator", "GetEnumerator"
    };
    public static ProjectionResult Project(DataGraph graph, WrapperContract contract, ProjectionPolicy policy, ProjectionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(graph);
        ArgumentNullException.ThrowIfNull(contract);
        ArgumentNullException.ThrowIfNull(policy);
        options ??= new ProjectionOptions();
        policy.ValidateOrThrow(options.ExpectedPolicyDigest);

        var dataValidation = DataGraphValidator.Validate(graph, options.ExpectedDataDigest, policy.DataSchemaVersion);
        if (!dataValidation.IsValid)
            throw new ProjectionValidationException(dataValidation.Issues.Select(static issue => new ProjectionIssue(issue.Code, issue.Message, issue.Path ?? "")).ToArray());
        ValidateContract(contract, policy.ContractSchemaVersion, options.ExpectedContractDigest);
        if (options.RejectUnresolvedAmbiguities && graph.Ambiguities.Any(static item => string.IsNullOrWhiteSpace(item.Resolution)))
            throw new ProjectionValidationException(graph.Ambiguities.Where(static item => string.IsNullOrWhiteSpace(item.Resolution)).Select(static item => new ProjectionIssue("data.ambiguity", item.Reason, item.LogicalId)).ToArray());
        if (graph.AbsentFacts.Count != 0)
            throw new ProjectionValidationException(graph.AbsentFacts.Select(static item => new ProjectionIssue("data.absent-fact", item.Reason, item.TargetId ?? item.LogicalId)).ToArray());

        var issues = new List<ProjectionIssue>();
        var libraryById = graph.Libraries.ToDictionary(static item => item.LogicalId, StringComparer.Ordinal);
        var projectById = graph.Projects.ToDictionary(static item => item.LogicalId, StringComparer.Ordinal);
        var typeById = graph.Types.ToDictionary(static item => item.LogicalId, StringComparer.Ordinal);
        var typeByName = graph.Types.GroupBy(static item => item.Name, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.ToArray(), StringComparer.Ordinal);
        var qualifiedTypeById = graph.Types.ToDictionary(static item => item.LogicalId, item => QualifyTypeName(item, libraryById, projectById), StringComparer.Ordinal);
        var dataTypeNamesByProduct = graph.Types
            .GroupBy(item => projectById.TryGetValue(item.ProjectId, out var ownerProject) ? ownerProject.Name : libraryById[item.LibraryId].Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(static group => group.Key, static group => group.Select(static item => item.Name).ToHashSet(StringComparer.Ordinal), StringComparer.OrdinalIgnoreCase);
        var membersByType = graph.Members.GroupBy(static item => item.TypeId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.OrderBy(item => item.LogicalId, StringComparer.Ordinal).ToArray(), StringComparer.Ordinal);
        var valuesByType = graph.Values.GroupBy(static item => item.TypeId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.OrderBy(item => item.LogicalId, StringComparer.Ordinal).ToArray(), StringComparer.Ordinal);
        var invocationById = graph.InvocationEvidence.ToDictionary(static item => item.LogicalId, StringComparer.Ordinal);
        var observationsByTarget = graph.SupportObservations.Where(static item => item.Present).GroupBy(static item => item.TargetId, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.ToArray(), StringComparer.Ordinal);
        var projectedDataMemberIds = new HashSet<string>(StringComparer.Ordinal);
        var projectedValueIds = new HashSet<string>(StringComparer.Ordinal);
        var files = new List<WrapperFile>();
        var traces = new List<ProjectionTrace>();

        foreach (var dataType in graph.Types.OrderBy(static item => item.LogicalId, StringComparer.Ordinal))
        {
            if (!libraryById.TryGetValue(dataType.LibraryId, out var library))
            {
                issues.Add(new("type.library", "Data type has no owning library.", dataType.LogicalId));
                continue;
            }

            projectById.TryGetValue(dataType.ProjectId, out var project);
            var product = project?.Name ?? library.Name;
            var descriptor = Describe(dataType);
            var contractType = FindContractType(dataType, product, contract, issues);
            var auxiliaryContractTypes = AuxiliaryContractTypes(contractType, contract, dataTypeNamesByProduct.TryGetValue(product, out var dataTypeNames) ? dataTypeNames : EmptyStringSet);
            var inheritedContractTypes = InheritedContractTypes(contractType, contract);
            var defaultIndexerNames = DefaultIndexerNames(contractType);
            var auxiliaryTypes = auxiliaryContractTypes.Select(contractType => ProjectAuxiliaryType(contractType, policy)).ToArray();
            var typeNamespace = contractType?.Namespace ?? NamespaceFromRoot(string.IsNullOrWhiteSpace(dataType.Namespace) ? "NetOffice." + product + "Api" : dataType.Namespace, descriptor.Category);
            var typeName = Sanitize(dataType.Name, policy.Names);
            var support = contractType is null
                ? SupportFor(dataType.LogicalId, observationsByTarget)
                : ParseSupport(contractType.Attributes);
            var attributes = ContractOrDataAttributes(contractType?.Attributes, support, descriptor.EntityKind);
            var sourceMembers = membersByType.GetValueOrDefault(dataType.LogicalId, Array.Empty<DataMember>());
            var capabilities = TypeCapabilities(dataType, sourceMembers, descriptor, contractType, policy);
            var runtime = TypeRuntimeRequirements(descriptor, capabilities, attributes, policy);
            var baseNames = ResolveBaseNames(dataType, qualifiedTypeById, contractType);
            var baseType = contractType?.BaseType ?? (contractType?.Kind.Equals("interface", StringComparison.OrdinalIgnoreCase) == true ? null : baseNames.FirstOrDefault() ?? DefaultBaseType(descriptor, capabilities));
            var interfaces = (contractType?.Interfaces ?? InterfaceNames(dataType, qualifiedTypeById, baseType)).Distinct(StringComparer.Ordinal).OrderBy(static item => item, StringComparer.Ordinal).ToArray();
            var effectiveBases = EffectiveBaseTypes(dataType, typeById).Select(type => QualifyTypeName(type, libraryById, projectById)).Distinct(StringComparer.Ordinal).ToList();
            if (baseType is not null && !effectiveBases.Contains(baseType, StringComparer.Ordinal)) effectiveBases.Insert(0, baseType);
            var constructors = ConstructorPlans(descriptor, product, dataType.Name, contractType, policy);
            var projectedMembers = new List<ProjectedMember>();

            foreach (var dataMember in sourceMembers)
            {
                var hasParameters = dataMember.Parameters.Count != 0 || dataMember.ParameterTypes.Count != 0;
                var isIndexer = dataMember.AccessorKind is not null && hasParameters
                                && (defaultIndexerNames.Count != 0
                                    ? defaultIndexerNames.Contains(dataMember.Name)
                                    : dataMember.Name.Equals("Item", StringComparison.Ordinal));
                var supersededByIndexer = defaultIndexerNames.Count != 0
                                          && !defaultIndexerNames.Contains("Item")
                                          && dataMember.Name.Equals("Item", StringComparison.Ordinal);
                var supersededByConstructor = dataMember.Name.Equals(typeName, StringComparison.Ordinal);
                var arities = OverloadArities(dataMember, isIndexer);
                var overloadOrdinal = 0;
                foreach (var arity in arities)
                {
                    var parameters = Parameters(dataMember, arity, policy.Names, qualifiedTypeById, isIndexer, descriptor.EntityKind == ProjectionEntityKind.EventInterface);
                    var memberMatch = ChooseContractMember(contractType,
                        descriptor.EntityKind == ProjectionEntityKind.EventInterface
                            ? auxiliaryContractTypes.Where(static candidate => !candidate.Name.EndsWith("_SinkHelper", StringComparison.Ordinal)).ToArray()
                            : auxiliaryContractTypes,
                        inheritedContractTypes, dataMember, arity, isIndexer);
                    var metadata = memberMatch.Member;
                    if (descriptor.EntityKind == ProjectionEntityKind.EventInterface && metadata is not null)
                        metadata = OverlayEventSinkInvocation(metadata, contract);
                    var disposition = supersededByConstructor
                        ? ProjectionMemberEmissionDisposition.SupersededByConstructorPlan
                        : supersededByIndexer ? ProjectionMemberEmissionDisposition.SupersededByIndexer
                        : metadata is null && contractType is not null ? ProjectionMemberEmissionDisposition.SuppressedByContract
                        : memberMatch.Disposition;
                    var memberSupport = metadata is null
                        ? SupportFor(dataMember.LogicalId, observationsByTarget)
                        : ParseSupport(metadata.Attributes);
                    var member = ProjectMember(dataMember, metadata, parameters, overloadOrdinal++, memberSupport, invocationById, typeByName, qualifiedTypeById, policy, memberMatch.OwnerLogicalId, disposition, isIndexer);
                    if (descriptor.EntityKind is ProjectionEntityKind.Module or ProjectionEntityKind.Constants)
                        member = member with { Modifiers = member.Modifiers.Concat(new[] { "static" }).Distinct(StringComparer.Ordinal).OrderBy(static item => item, StringComparer.Ordinal).ToArray() };
                    projectedMembers.Add(member);
                    projectedDataMemberIds.Add(dataMember.LogicalId);
                    if (options.IncludeTraces) traces.Add(MemberTrace(member, dataMember, metadata, arity));
                }
            }

            if (contractType is not null)
            {
                foreach (var contractEvent in contractType.Members.Where(static member => member.Kind.Equals("event", StringComparison.OrdinalIgnoreCase))
                             .Where(contractEvent => projectedMembers.All(member => !member.CSharpName.Equals(contractEvent.Name, StringComparison.Ordinal)))
                             .OrderBy(static member => member.Name, StringComparer.Ordinal))
                    projectedMembers.Add(ProjectContractEvent(contractEvent, contractType.LogicalId, policy));

                var projectedContractSignatures = projectedMembers.Select(static member => member.EmissionOwnerLogicalId + "\u001f" + member.Signature).ToHashSet(StringComparer.Ordinal);
                foreach (var contractOwner in new[] { contractType }.Concat(auxiliaryContractTypes).OrderBy(static type => type.LogicalId, StringComparer.Ordinal))
                {
                    foreach (var contractMember in contractOwner.Members
                                 .Where(static member => member.Kind is not "constructor" and not "event" and not "field" and not "enum-value")
                                 .Where(static member => !IsCollectionInfrastructureMember(member))
                                 .Where(static member => !IsRuntimeInfrastructureMember(member))
                                 .Where(member => !contractOwner.Name.EndsWith("_SinkHelper", StringComparison.Ordinal)
                                                  || !projectedMembers.Any(projected =>
                                                      projected.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit
                                                      && projected.EmissionOwnerLogicalId.Equals(contractType.LogicalId, StringComparison.Ordinal)
                                                      && projected.CSharpName.Equals(Sanitize(member.Name, policy.Names), StringComparison.Ordinal)
                                                      && string.Equals(projected.Parameters, member.Parameters, StringComparison.Ordinal)))
                                 .Where(member => !projectedContractSignatures.Contains(contractOwner.LogicalId + "\u001f" + member.Signature))
                                 .OrderBy(static member => member.Name, StringComparer.Ordinal)
                                 .ThenBy(static member => member.Signature, StringComparer.Ordinal)
                                 .ThenBy(static member => member.Source, StringComparer.Ordinal)
                                 .ThenBy(static member => member.Line))
                    {
                        projectedMembers.Add(ProjectContractModuleMember(contractMember, contractOwner.LogicalId, projectedMembers.Count, typeByName, policy, false));
                        projectedContractSignatures.Add(contractOwner.LogicalId + "\u001f" + contractMember.Signature);
                    }
                }
            }
            foreach (var value in valuesByType.GetValueOrDefault(dataType.LogicalId, Array.Empty<DataValue>()))
            {
                var valueMetadata = contractType?.Members.Where(member => member.Name.Equals(value.Name, StringComparison.Ordinal))
                    .OrderBy(static member => member.Source, StringComparer.Ordinal).ThenBy(static member => member.Line).FirstOrDefault();
                var valueSupport = valueMetadata is null ? SupportFor(value.LogicalId, observationsByTarget) : ParseSupport(valueMetadata.Attributes);
                var valueMember = ProjectValue(value, descriptor, valueSupport, policy, contractType?.LogicalId ?? dataType.LogicalId, valueMetadata);
                if (valueMetadata is null && contractType is not null)
                    valueMember = valueMember with { EmissionDisposition = ProjectionMemberEmissionDisposition.SuppressedByContract };
                if (projectedMembers.Any(member => member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit && member.CSharpName.Equals(valueMember.CSharpName, StringComparison.Ordinal)))
                    valueMember = valueMember with { EmissionDisposition = ProjectionMemberEmissionDisposition.SuppressedByContract };
                projectedMembers.Add(valueMember);
                projectedValueIds.Add(value.LogicalId);
                if (options.IncludeTraces)
                    traces.Add(new ProjectionTrace
                {
                    LogicalId = valueMember.LogicalId,
                    Steps = new[]
                    {
                        Step("source", "enumerate-data-v2-values", value.LogicalId),
                        Step("kind", "map-value-category", valueMember.Kind),
                        Step("value", "preserve-canonical-value", value.Value),
                        Step("docs", "stable-logical-binding-key", valueMember.DocsBindingKey)
                    }
                });
            }

            projectedMembers = MarkDuplicates(projectedMembers.OrderBy(static item => item.Accessors.Contains("get", StringComparer.Ordinal) ? 0 : item.Accessors.Contains("set", StringComparer.Ordinal) ? 1 : 2).ThenBy(static item => item.DataLogicalId, StringComparer.Ordinal).ThenBy(static item => item.OverloadOrdinal).ThenBy(static item => item.LogicalId, StringComparer.Ordinal).ToList());
            var propertyNames = projectedMembers.Where(static member => member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit && member.Kind is "property" or "indexer").Select(static member => member.EmissionOwnerLogicalId + "\u001f" + member.CSharpName).ToHashSet(StringComparer.Ordinal);
            projectedMembers = projectedMembers.Select(member => member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit
                                                                 && member.Kind == "method"
                                                                 && propertyNames.Contains(member.EmissionOwnerLogicalId + "\u001f" + member.CSharpName)
                ? member with { EmissionDisposition = ProjectionMemberEmissionDisposition.SupersededByProperty }
                : member).ToList();
            var duplicateGroups = projectedMembers.Where(static item => item.DuplicateOf is not null).Select(static item => item.OverloadGroup).Distinct(StringComparer.Ordinal).OrderBy(static item => item, StringComparer.Ordinal).ToArray();
            var source = contractType is null ? DataLocation(dataType.Provenance) : ContractPrimarySource(contractType);
            var projectedType = new ProjectedType
            {
                LogicalId = contractType?.LogicalId ?? dataType.LogicalId,
                CanonicalLogicalId = contractType?.LogicalId ?? dataType.LogicalId,
                ContractPartLogicalId = contractType?.LogicalId,
                ContractPartSource = source,
                DataLogicalId = dataType.LogicalId,
                LibraryLogicalId = library.LogicalId,
                Product = product,
                Namespace = typeNamespace,
                Name = dataType.Name,
                CSharpName = typeName,
                Kind = contractType?.Kind ?? descriptor.CSharpKind,
                EntityKind = descriptor.EntityKind,
                EmissionDisposition = contractType is null && product.Equals(contract.Source.Api, StringComparison.OrdinalIgnoreCase)
                    ? ProjectionEmissionDisposition.SuppressedByContract
                    : IsCompanionContractType(contractType) ? ProjectionEmissionDisposition.Companion : ProjectionEmissionDisposition.Emit,
                CompanionPath = IsCompanionContractType(contractType) ? contractType!.Source : null,
                FileCategory = descriptor.Category,
                Accessibility = NormalizeTopLevelAccessibility(contractType?.Accessibility),
                Modifiers = TypeModifiers(descriptor, contractType),
                Attributes = attributes,
                DeclaredGuid = dataType.DeclaredGuid,
                TypeLibType = dataType.TypeLibType,
                IsHidden = dataType.IsHidden == true,
                IsEarlyBind = dataType.IsEarlyBind == true,
                AutomaticQuit = dataType.AutomaticQuit == true,
                IsApplicationObject = dataType.IsApplicationObject == true,
                BaseType = baseType,
                Interfaces = interfaces,
                AliasTarget = dataType.AliasTarget,
                ProgId = descriptor.EntityKind == ProjectionEntityKind.CoClass ? ActivationProgId(constructors, product + "." + dataType.Name) : null,
                EventSink = descriptor.EntityKind == ProjectionEntityKind.EventInterface
                    ? new EventSinkPlan { Name = typeName + "_SinkHelper", InterfaceName = typeName, InterfaceId = dataType.DeclaredGuid }
                    : null,
                EventBindings = descriptor.EntityKind == ProjectionEntityKind.CoClass
                    ? ProjectEventBindings(dataType, contractType, contract, typeById, qualifiedTypeById)
                    : Array.Empty<ProjectedEventBinding>(),
                EffectiveBaseTypes = effectiveBases,
                DuplicateGroups = duplicateGroups,
                AuxiliaryTypes = auxiliaryTypes,
                Signature = contractType?.Signature ?? TypeSignature(descriptor, typeName, baseType, interfaces),
                SupportVersions = support,
                Capabilities = capabilities,
                RuntimeRequirements = runtime,
                Constructors = constructors,
                ExactContractMembers = ProjectExactContractMembers(contractType, policy, projectedMembers),
                DocsBindingKey = DocsKey(policy, dataType.LogicalId),
                Source = source,
                Line = contractType?.Line ?? SourceLine(dataType.Provenance),
                Members = projectedMembers
            };
            projectedType = ApplyTypeOverrides(projectedType, policy, issues);
            var partitionedFiles = PartitionContractParts(projectedType, contract, source, policy);
            files.AddRange(partitionedFiles);
            if (options.IncludeTraces) traces.Add(TypeTrace(projectedType, dataType, contractType, partitionedFiles[0].Path));
        }

        foreach (var contractModule in contract.Types
                     .Where(static type => type.Attributes.Any(static attribute => attribute.Contains("EntityType(EntityType.IsModule)", StringComparison.Ordinal)))
                     .Where(type => files.SelectMany(static file => file.Types).All(projected => !projected.LogicalId.Equals(type.LogicalId, StringComparison.Ordinal)))
                     .OrderBy(static type => type.LogicalId, StringComparer.Ordinal))
        {
            var module = ProjectContractModule(contractModule, contract.Source.Api, typeByName, policy);
            files.Add(new WrapperFile
            {
                Path = PartitionPath(module, contractModule.Source, policy),
                Namespace = module.Namespace,
                Category = module.FileCategory,
                Types = new[] { module }
            });
        }

        foreach (var projectInfoContract in contract.Types
                     .Where(static type => type.Source.Equals("Utils/ProjectInfo.cs", StringComparison.OrdinalIgnoreCase))
                     .Where(type => files.SelectMany(static file => file.Types).All(projected => !projected.LogicalId.Equals(type.LogicalId, StringComparison.Ordinal)))
                     .OrderBy(static type => type.LogicalId, StringComparer.Ordinal))
        {
            var ownerProject = graph.Projects.SingleOrDefault(project => project.Name.Equals(contract.Source.Api, StringComparison.OrdinalIgnoreCase))
                ?? throw new ProjectionValidationException(new[] { new ProjectionIssue("project-info.project", "ProjectInfo contract has no matching Data v2 project.", projectInfoContract.LogicalId) });
            var projectInfo = ProjectContractProjectInfo(projectInfoContract, ownerProject, libraryById, projectById, policy);
            files.Add(new WrapperFile
            {
                Path = PartitionPath(projectInfo, projectInfoContract.Source, policy),
                Namespace = projectInfo.Namespace,
                Category = projectInfo.FileCategory,
                Types = new[] { projectInfo }
            });
            if (options.IncludeTraces)
                traces.Add(new ProjectionTrace
            {
                LogicalId = projectInfo.LogicalId,
                Steps = new[]
                {
                    Step("source", "bind-project-info-contract", projectInfoContract.LogicalId),
                    Step("data", "derive-project-info-metadata", ownerProject.LogicalId),
                    Step("partition", "preserve-contract-source", projectInfoContract.Source)
                }
            });
        }

        ApplyMemberOverrides(files, policy, issues);
        if (options.IncludeTraces) AppendOverrideTraces(files, policy, traces);
        ValidateCoverage(graph, files, projectedDataMemberIds, projectedValueIds, issues);
        if (issues.Count != 0) throw new ProjectionValidationException(issues);

        return new ProjectionResult
        {
            DataSchemaVersion = graph.SchemaVersion,
            DataDigest = graph.Digest,
            ContractSchemaVersion = contract.SchemaVersion,
            ContractDigest = ContractDigest(contract),
            PolicyDigest = policy.Digest,
            Files = files.OrderBy(static file => file.Path, StringComparer.Ordinal).ThenBy(static file => file.Namespace, StringComparer.Ordinal).ToArray(),
            Explain = options.IncludeTraces ? traces.OrderBy(static trace => trace.LogicalId, StringComparer.Ordinal).ToArray() : Array.Empty<ProjectionTrace>(),
            Coverage = new ProjectionCoverage
            {
                DataTypes = graph.Types.Count,
                ProjectedTypes = files.Sum(static file => file.Types.Count),
                DataMembers = graph.Members.Count,
                ProjectedDataMembers = projectedDataMemberIds.Count,
                DataValues = graph.Values.Count,
                ProjectedValues = projectedValueIds.Count
            }
        };
    }

    public static string ProjectJson(string dataGraphJson, string contractJson, string policyJson, ProjectionOptions? options = null, bool indented = false)
        => Project(CanonicalJson.Parse(dataGraphJson), WrapperContract.Parse(contractJson), ProjectionPolicy.Parse(policyJson), options).ToJson(indented);

    public static string Serialize(ProjectionResult result, bool indented = false) => result.ToJson(indented);

    private static TypeDescriptor Describe(DataType type)
    {
        if (type.IsEventInterface == true) return new(ProjectionEntityKind.EventInterface, ProjectionFileCategory.Events, "interface");
        var source = DataLocation(type.Provenance).Replace('\\', '/');
        var descriptor = type.Kind.Trim().ToLowerInvariant() switch
        {
            "dispatch" or "dispatchinterface" or "dual" => new TypeDescriptor(ProjectionEntityKind.DispatchInterface, ProjectionFileCategory.DispatchInterfaces, "class"),
            "interface" when source.Contains("/DispatchInterfaces.xml", StringComparison.OrdinalIgnoreCase) || source.StartsWith("DispatchInterfaces.xml", StringComparison.OrdinalIgnoreCase) => new TypeDescriptor(ProjectionEntityKind.DispatchInterface, ProjectionFileCategory.DispatchInterfaces, "class"),
            "interface" => new TypeDescriptor(ProjectionEntityKind.Interface, ProjectionFileCategory.Interfaces, "class"),
            "class" or "coclass" or "classmodule" => new TypeDescriptor(ProjectionEntityKind.CoClass, ProjectionFileCategory.Classes, "class"),
            "enum" => new TypeDescriptor(ProjectionEntityKind.Enum, ProjectionFileCategory.Enums, "enum"),
            "module" => new TypeDescriptor(ProjectionEntityKind.Module, ProjectionFileCategory.Modules, "class"),
            "constant" or "constants" => new TypeDescriptor(ProjectionEntityKind.Constants, ProjectionFileCategory.Constants, "class"),
            "record" or "struct" or "recordstruct" => new TypeDescriptor(ProjectionEntityKind.Record, ProjectionFileCategory.Records, "struct"),
            "alias" or "typedef" => new TypeDescriptor(ProjectionEntityKind.TypeDef, ProjectionFileCategory.TypeDefs, "struct"),
            _ => throw new ProjectionValidationException(new[] { new ProjectionIssue("type.kind.unsupported", $"Unsupported Data v2 type kind '{type.Kind}'.", type.LogicalId) })
        };
        return descriptor with { Category = SourceCategory(type.SourceCategory, descriptor.Category) };
    }

    private static string ContractPrimarySource(ContractType contract)
    {
        var canonicalFileName = contract.Name + ".cs";
        return new[] { contract.Source }.Concat(contract.Members.Select(static member => member.Source))
                   .Where(static source => !string.IsNullOrWhiteSpace(source))
                   .Distinct(StringComparer.Ordinal)
                   .Where(source => Path.GetFileName(source).Equals(canonicalFileName, StringComparison.OrdinalIgnoreCase))
                   .OrderBy(static source => source.Length)
                   .ThenBy(static source => source, StringComparer.Ordinal)
                   .FirstOrDefault()
               ?? contract.Source;
    }

    private static ContractType? FindContractType(DataType type, string product, WrapperContract contract, ICollection<ProjectionIssue> issues)
    {
        if (!string.Equals(contract.Source.Api, product, StringComparison.OrdinalIgnoreCase)) return null;

        var matches = contract.Types
            .Where(candidate => string.Equals(candidate.Name, type.Name, StringComparison.Ordinal))
            .OrderBy(static candidate => candidate.LogicalId, StringComparer.Ordinal)
            .ThenBy(static candidate => candidate.Source, StringComparer.Ordinal)
            .ToArray();
        if (matches.Length <= 1) return matches.FirstOrDefault();

        var sourceCategory = NormalizeSourceCategory(type.SourceCategory);
        var categoryMatches = matches.Where(candidate => string.Equals(NormalizeSourceCategory(ContractSourceCategory(candidate.Source)), sourceCategory, StringComparison.Ordinal)).ToArray();
        if (categoryMatches.Length != 0)
        {
            matches = categoryMatches;
            matches = PreferMatches(matches, candidate => string.Equals(NormalizeContractNamespace(candidate), type.Namespace, StringComparison.Ordinal));
        }
        else
        {
            matches = PreferMatches(matches, candidate => string.Equals(candidate.Namespace, type.Namespace, StringComparison.Ordinal));
        }
        if (matches.Length <= 1) return matches.FirstOrDefault();

        if (CanMergeContractFragments(matches))
            return MergeContractFragments(matches);

        issues.Add(new("type.conflict", $"Wrapper Contract has {matches.Length} incompatible identities named {type.Name} for {product} after namespace and source-category matching.", type.LogicalId));
        return null;
    }

    private static ContractType[] PreferMatches(ContractType[] candidates, Func<ContractType, bool> predicate)
    {
        var preferred = candidates.Where(predicate).ToArray();
        return preferred.Length == 0 ? candidates : preferred;
    }

    private static bool CanMergeContractFragments(IReadOnlyList<ContractType> candidates)
    {
        var first = candidates[0];
        var normalizedNamespace = NormalizeContractNamespace(first);
        return candidates.All(candidate =>
            string.Equals(NormalizeContractNamespace(candidate), normalizedNamespace, StringComparison.Ordinal)
            && string.Equals(candidate.Name, first.Name, StringComparison.Ordinal)
            && string.Equals(candidate.Kind, first.Kind, StringComparison.Ordinal)
            && string.Equals(candidate.Accessibility, first.Accessibility, StringComparison.Ordinal)
            && string.Equals(candidate.BaseType, first.BaseType, StringComparison.Ordinal)
            && string.Equals(NormalizeTypeSignature(candidate.Signature), NormalizeTypeSignature(first.Signature), StringComparison.Ordinal));
    }

    private static ContractType MergeContractFragments(IReadOnlyList<ContractType> candidates)
    {
        var expectedLogicalId = NormalizeContractNamespace(candidates[0]) + "." + candidates[0].Name;
        var primary = candidates
            .OrderByDescending(candidate => Path.GetFileName(candidate.Source).Equals(candidate.Name + ".cs", StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(candidate => string.Equals(candidate.LogicalId, expectedLogicalId, StringComparison.Ordinal))
            .ThenByDescending(static candidate => candidate.Members.Count)
            .ThenBy(static candidate => candidate.LogicalId, StringComparer.Ordinal)
            .ThenBy(static candidate => candidate.Source, StringComparer.Ordinal)
            .First();
        var members = candidates
            .SelectMany(static candidate => candidate.Members)
            .GroupBy(ContractMemberIdentity, StringComparer.Ordinal)
            .Select(static group => group.OrderBy(static member => member.Source, StringComparer.Ordinal).ThenBy(static member => member.Line).First())
            .OrderBy(static member => member.Name, StringComparer.Ordinal)
            .ThenBy(static member => member.Signature, StringComparer.Ordinal)
            .ThenBy(static member => member.Source, StringComparer.Ordinal)
            .ThenBy(static member => member.Line)
            .ToArray();

        return primary with
        {
            Modifiers = candidates.SelectMany(static candidate => candidate.Modifiers).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
            Interfaces = candidates.SelectMany(static candidate => candidate.Interfaces).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
            Attributes = candidates.SelectMany(static candidate => candidate.Attributes).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
            Members = members
        };
    }

    private static string ContractMemberIdentity(ContractMember member)
        => string.Join("\u001f", member.Name, member.Kind, member.Accessibility, member.ReturnType ?? "", member.Parameters ?? "", member.Signature);

    private static string NormalizeTypeSignature(string signature)
        => signature.Replace("partial ", "", StringComparison.Ordinal).Trim();

    private static string ContractSourceCategory(string source)
    {
        var normalized = source.Replace('\\', '/').TrimStart('/');
        var separator = normalized.IndexOf('/');
        return separator < 0 ? "" : normalized[..separator];
    }

    private static string NormalizeContractNamespace(ContractType candidate)
    {
        var category = ContractSourceCategory(candidate.Source);
        var suffix = "." + category;
        return category.Length != 0 && candidate.Namespace.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            ? candidate.Namespace[..^suffix.Length]
            : candidate.Namespace;
    }
    private static IReadOnlyList<ContractType> AuxiliaryContractTypes(ContractType? primary, WrapperContract contract, IReadOnlySet<string> dataTypeNames)
    {
        if (primary is null || string.IsNullOrWhiteSpace(primary.Source)) return Array.Empty<ContractType>();
        return contract.Types
            .Where(candidate => !string.Equals(candidate.Name, primary.Name, StringComparison.Ordinal)
                                && string.Equals(candidate.Source, primary.Source, StringComparison.Ordinal)
                                && !dataTypeNames.Contains(candidate.Name))
            .OrderBy(static candidate => candidate.LogicalId, StringComparer.Ordinal)
            .ToArray();
    }

    private static ProjectedAuxiliaryType ProjectAuxiliaryType(ContractType contractType, ProjectionPolicy policy)
    {
        var delegateMatch = DelegateRegex.Match(contractType.Signature);
        return new ProjectedAuxiliaryType
        {
            LogicalId = contractType.LogicalId,
            Namespace = contractType.Namespace,
            Name = contractType.Name,
            Kind = contractType.Kind,
            Accessibility = contractType.Accessibility,
            Modifiers = contractType.Modifiers.OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
            BaseType = contractType.BaseType,
            Interfaces = contractType.Interfaces.OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
            Attributes = contractType.Attributes.Select(NormalizeAttribute).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
            Signature = contractType.Signature,
            Source = contractType.Source,
            Line = contractType.Line,
            DelegateReturnType = delegateMatch.Success ? delegateMatch.Groups["return"].Value : null,
            DelegateParameters = delegateMatch.Success ? ParseDelegateParameters(delegateMatch.Groups["parameters"].Value) : Array.Empty<ProjectedParameter>(),
            Constructors = contractType.Kind.Equals("class", StringComparison.OrdinalIgnoreCase)
                ? ConstructorPlans(new TypeDescriptor(ProjectionEntityKind.DispatchInterface, ProjectionFileCategory.DispatchInterfaces, "class"), "", contractType.Name, contractType, policy)
                : Array.Empty<ConstructorPlan>()
        };
    }

    private static IReadOnlyList<ProjectedParameter> ParseDelegateParameters(string parameters)
    {
        if (string.IsNullOrWhiteSpace(parameters)) return Array.Empty<ProjectedParameter>();
        return SplitTopLevel(parameters).Select((text, index) => ParseParameterDeclaration(text, index, false)).ToArray();
    }

    private static ProjectedParameter ParseParameterDeclaration(string text, int index, bool contractOverlay)
    {
        var declaration = text.Trim();
        var working = declaration;
        var attributes = new List<string>();
        while (working.StartsWith("[", StringComparison.Ordinal))
        {
            var end = working.IndexOf(']');
            if (end < 0) break;
            attributes.Add(working[1..end].Trim());
            working = working[(end + 1)..].TrimStart();
        }
        var equals = IndexOfTopLevel(working, '=');
        var defaultValue = equals < 0 ? null : working[(equals + 1)..].Trim();
        if (equals >= 0) working = working[..equals].TrimEnd();
        var separator = working.LastIndexOf(' ');
        var name = separator < 0 ? "arg" + index.ToString(CultureInfo.InvariantCulture) : working[(separator + 1)..].Trim();
        var type = separator < 0 ? working : working[..separator].Trim();
        var refKind = "value";
        foreach (var modifier in new[] { "ref ", "out ", "in " })
            if (type.StartsWith(modifier, StringComparison.Ordinal))
            {
                refKind = modifier.Trim();
                type = type[modifier.Length..].Trim();
                break;
            }
        return new ProjectedParameter
        {
            Name = name,
            CSharpName = name,
            Type = QualifyBclType(type),
            RefKind = refKind,
            IsOptional = defaultValue is not null,
            HasDefaultValue = defaultValue is not null,
            DefaultValue = defaultValue,
            EmitDefaultValue = defaultValue is not null,
            Position = index,
            IsContractOverlay = contractOverlay,
            ContractDeclaration = contractOverlay ? declaration : null,
            Attributes = attributes,
            PreserveContractAttributes = contractOverlay
        };
    }

    private static int IndexOfTopLevel(string value, char target)
    {
        var depth = 0;
        for (var index = 0; index < value.Length; index++)
        {
            depth += value[index] is '<' or '(' or '[' or '{' ? 1 : value[index] is '>' or ')' or ']' or '}' ? -1 : 0;
            if (value[index] == target && depth == 0) return index;
        }
        return -1;
    }

    private static IReadOnlyList<string> SplitTopLevel(string value)
    {
        var result = new List<string>();
        var start = 0;
        var depth = 0;
        for (var index = 0; index < value.Length; index++)
        {
            depth += value[index] is '<' or '(' or '[' or '{' ? 1 : value[index] is '>' or ')' or ']' or '}' ? -1 : 0;
            if (value[index] != ',' || depth != 0) continue;
            result.Add(value[start..index]);
            start = index + 1;
        }
        result.Add(value[start..]);
        return result;
    }

    private static IReadOnlyList<ContractType> InheritedContractTypes(ContractType? primary, WrapperContract contract)
    {
        if (primary is null) return Array.Empty<ContractType>();
        var result = new List<ContractType>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Queue<string>(new[] { primary.BaseType }.Where(static name => !string.IsNullOrWhiteSpace(name)).Cast<string>().Concat(primary.Interfaces));
        while (pending.Count != 0)
        {
            var requestedName = pending.Dequeue().Split('.').Last();
            var candidates = contract.Types.Where(candidate => candidate.Name.Equals(requestedName, StringComparison.Ordinal)).ToArray();
            if (candidates.Length == 0) continue;
            var inherited = candidates.FirstOrDefault(candidate => candidate.Namespace.Equals(primary.Namespace, StringComparison.Ordinal))
                            ?? candidates.OrderBy(static candidate => candidate.LogicalId, StringComparer.Ordinal).First();
            if (!seen.Add(inherited.LogicalId)) continue;
            result.Add(inherited);
            foreach (var sibling in contract.Types.Where(candidate => !candidate.Name.Equals(inherited.Name, StringComparison.Ordinal)
                                                                      && candidate.Source.Equals(inherited.Source, StringComparison.Ordinal))
                                                    .OrderBy(static candidate => candidate.LogicalId, StringComparer.Ordinal))
                if (seen.Add(sibling.LogicalId)) result.Add(sibling);
            if (!string.IsNullOrWhiteSpace(inherited.BaseType)) pending.Enqueue(inherited.BaseType);
            foreach (var inheritedInterface in inherited.Interfaces) pending.Enqueue(inheritedInterface);
        }
        return result;
    }

    private static IReadOnlySet<string> DefaultIndexerNames(ContractType? contractType)
    {
        if (contractType is null) return EmptyStringSet;
        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var attribute in contractType.Attributes)
        {
            var match = IndexPropertyRegex.Match(attribute);
            if (match.Success) result.Add(match.Groups["name"].Value);
        }
        return result;
    }


    private static bool IsCompanionContractType(ContractType? contractType)
    {
        if (contractType is null) return false;
        var category = NormalizeSourceCategory(ContractSourceCategory(contractType.Source));
        return category is "native" or "nativecaller";
    }

    private static string NormalizeSourceCategory(string category)
        => category.Trim().ToLowerInvariant() switch
        {
            "coclasses" => "classes",
            "dispatchinterfaces" => "dispatchinterfaces",
            "interfaces" => "interfaces",
            "events" => "events",
            "classes" => "classes",
            "enums" => "enums",
            "modules" => "modules",
            "constants" => "constants",
            "records" => "records",
            "typedefs" => "typedefs",
            _ => category.Trim().ToLowerInvariant()
        };

    private static ProjectedMember ProjectMember(DataMember data, ContractMember? metadata, IReadOnlyList<ProjectedParameter> parameters, int overloadOrdinal, IReadOnlyList<VersionSupport> support, IReadOnlyDictionary<string, InvocationEvidence> invocationById, IReadOnlyDictionary<string, DataType[]> typeByName, IReadOnlyDictionary<string, string> qualifiedTypeById, ProjectionPolicy policy, string ownerLogicalId, ProjectionMemberEmissionDisposition disposition, bool isIndexer)
    {
        var name = metadata?.Name ?? data.Name;
        var csharpName = Sanitize(name, policy.Names);
        var parsedContractParameters = metadata is not null && ParameterCount(metadata.Parameters) == parameters.Count
            ? ParseContractParameters(metadata.Parameters)
            : Array.Empty<ProjectedParameter>();
        var contractParameters = parsedContractParameters.Select((parameter, index) => parameter with
        {
            TypeReference = null,
            SinkArgument = OverlaySinkArgument(parameters[index].SinkArgument, parameter)
        }).ToArray();
        contractParameters = ApplyContractSinkArguments(contractParameters, metadata?.Attributes);
        contractParameters = ApplyContractSinkInvocations(contractParameters, metadata?.InvocationText);
        var effectiveParameters = contractParameters.Length == parameters.Count ? contractParameters : parameters;
        var kind = NormalizeMemberKind(metadata?.Kind ?? data.Kind, data.AccessorKind, effectiveParameters.Count, isIndexer);
        var parameterText = metadata?.Parameters;
        if (parameterText is null || ParameterCount(parameterText) != effectiveParameters.Count) parameterText = FormatParameters(effectiveParameters);
        var qualifiedDataReturnType = metadata?.ReturnType is { Length: > 0 } contractReturnType
            ? contractReturnType
            : QualifiedReferenceType(data.ReturnTypeReference, data.ReturnType ?? (kind is "property" or "indexer" ? "object" : "void"), qualifiedTypeById);
        var nativeResult = metadata?.Attributes.Any(static attribute => attribute.Contains("NativeResult", StringComparison.Ordinal)) == true
                           || data.ReturnTypeReference?.IsExternal == true && qualifiedDataReturnType.StartsWith("stdole.", StringComparison.OrdinalIgnoreCase);
        var returnType = QualifyBclType(nativeResult ? metadata?.ReturnType ?? qualifiedDataReturnType : qualifiedDataReturnType);
        var attributes = metadata is null ? DataMemberAttributes(null, support, data) : ContractOrSupportAttributes(metadata.Attributes, support);
        invocationById.TryGetValue(data.InvocationEvidenceId ?? "", out var evidence);
        var invocation = Invocation(data, evidence, metadata, effectiveParameters, returnType, kind, typeByName, qualifiedTypeById);
        var primaryInvocation = OverlayContractInvocation(nativeResult ? invocation with { ReturnConversion = ProjectionReturnConversion.Native, ResultType = returnType } : invocation, metadata, effectiveParameters, returnType, typeByName);
        var accessorInvocations = data.AccessorKind is null ? new List<InvocationPlan>() : new List<InvocationPlan> { primaryInvocation };
        var accessors = data.AccessorKind is null ? new List<string>() : new List<string> { data.AccessorKind == "get" ? "get" : "set" };
        if (kind is "property" or "indexer"
            && !accessors.Contains("get", StringComparer.Ordinal)
            && metadata?.InvocationText.Any(static text => text.Contains("PropertyGet", StringComparison.Ordinal)) == true)
        {
            var getterText = metadata.InvocationText.Where(static text => text.Contains("PropertyGet", StringComparison.Ordinal)).ToArray();
            var getterFallback = nativeResult ? ProjectionReturnConversion.Native : ReturnConversion(returnType, null, typeByName);
            var getterConversion = ContractReturnConversion(string.Join("\n", getterText), metadata, returnType, typeByName, getterFallback);
            var getterArguments = InvocationArguments(effectiveParameters, false, returnType);
            var getterLocalReturn = ContractLocalReturn(getterText);
            accessorInvocations.Insert(0, primaryInvocation with
            {
                Operation = "property-get",
                OperationKind = ProjectionInvocationOperation.PropertyGet,
                Api = ContractInvocationApi(getterText, getterArguments.Any(static argument => argument.ByRef) ? ProjectionInvocationApi.Invoker : ProjectionInvocationApi.Factory),
                ArgumentCount = effectiveParameters.Count,
                ArgumentOrder = effectiveParameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName).ToArray(),
                ResultType = returnType,
                ResultTypeReference = null,
                SetValueConversion = null,
                ReturnConversion = getterConversion,
                FactoryMethodSuffix = ContractFactoryMethodSuffix(getterText) ?? ScalarFactorySuffix(returnType),
                RequiresProxy = getterConversion is ProjectionReturnConversion.KnownReference or ProjectionReturnConversion.Reference or ProjectionReturnConversion.UntypedReference or ProjectionReturnConversion.BaseReference,
                ArgumentPacking = ContractArgumentPacking(getterText),
                ObjectArrayStyle = ContractObjectArrayStyle(getterText),
                ReturnCastStyle = ContractReturnCastStyle(getterText),
                RawCallCast = ContractRawCallCast(getterText),
                KnownReferenceFactoryStyle = ContractKnownReferenceFactoryStyle(getterText),
                ReturnValueStyle = getterLocalReturn.Name is null ? ProjectionReturnValueStyle.Direct : ProjectionReturnValueStyle.Local,
                LocalReturnName = getterLocalReturn.Name,
                LocalReturnType = getterLocalReturn.Type,
                InvokerCallStyle = ContractInvokerCallStyle(getterText),
                CallKind = ContractInvocationCallKind(getterText),
                HasContractInvocation = getterText.Length != 0,
                Arguments = OverlayContractArgumentExpressions(getterArguments, getterText),
                ReleaseArguments = ContractReleaseArguments(getterText)
            });
            accessors.Insert(0, "get");
        }
        var capabilities = MemberCapabilities(data, metadata, kind, primaryInvocation, policy);
        var logicalId = ownerLogicalId + "/" + data.LogicalId + (overloadOrdinal == 0 ? "" : "/overload/" + effectiveParameters.Count.ToString(CultureInfo.InvariantCulture));
        return new ProjectedMember
        {
            LogicalId = logicalId,
            DataLogicalId = data.LogicalId,
            EmissionOwnerLogicalId = ownerLogicalId,
            EmissionDisposition = disposition,
            Name = name,
            CSharpName = csharpName,
            Kind = kind,
            Accessibility = metadata?.Accessibility ?? "public",
            Modifiers = (metadata?.Modifiers ?? Array.Empty<string>()).OrderBy(static item => item, StringComparer.Ordinal).ToArray(),
            ReturnType = returnType,
            Parameters = parameterText,
            ParameterList = effectiveParameters,
            Signature = metadata?.Signature is { Length: > 0 } signature && ParameterCount(metadata.Parameters) == effectiveParameters.Count ? signature : DataSignature(kind, data.AccessorKind, returnType, csharpName, parameterText),
            OverloadGroup = ownerLogicalId + ":overload:" + csharpName,
            OverloadOrdinal = overloadOrdinal,
            DuplicateOf = null,
            Invocation = primaryInvocation,
            AccessorInvocations = accessorInvocations,
            SupportVersions = support,
            Capabilities = capabilities,
            RuntimeRequirements = RuntimeRequirements("member", kind, attributes, policy).Concat(capabilities.Select(static capability => "capability:" + capability)).Distinct(StringComparer.Ordinal).OrderBy(static item => item, StringComparer.Ordinal).ToArray(),
            IsHidden = data.IsHidden,
            AnalyzeReturn = data.AnalyzeReturn,
            IsComProxy = data.IsComProxy,
            DocsBindingKey = DocsKey(policy, data.LogicalId),
            Attributes = attributes,
            AccessorGroupId = data.AccessorGroupId,
            Accessors = accessors,
            UseCustomEventAccessors = HasCustomEventAccessors(metadata),
            EventBackingField = ContractEventBackingField(metadata),
            EventValidationKey = ContractEventValidationKey(metadata?.InvocationText),
            EventInvalidReleaseArguments = ContractEventInvalidReleaseArguments(metadata?.InvocationText, effectiveParameters),
            EventValidationInlineReturn = ContractEventValidationInlineReturn(metadata?.InvocationText),
            Source = metadata?.Source ?? DataLocation(data.Provenance),
            Line = metadata?.Line ?? SourceLine(data.Provenance)
        };
    }

    private static ProjectedMember ProjectContractEvent(ContractMember metadata, string ownerLogicalId, ProjectionPolicy policy)
    {
        var returnType = QualifyBclType(metadata.ReturnType ?? "System.EventHandler");
        var support = ParseSupport(metadata.Attributes);
        var attributes = metadata.Attributes.Select(NormalizeAttribute).Distinct(StringComparer.Ordinal).OrderBy(static attribute => attribute, StringComparer.Ordinal).ToArray();
        return new ProjectedMember
        {
            LogicalId = ownerLogicalId + "/contract-event/" + metadata.Name,
            DataLogicalId = ownerLogicalId + "/contract-event/" + metadata.Name,
            EmissionOwnerLogicalId = ownerLogicalId,
            IsContractDerived = true,
            Name = metadata.Name,
            CSharpName = Sanitize(metadata.Name, policy.Names),
            Kind = "event",
            Accessibility = metadata.Accessibility,
            Modifiers = metadata.Modifiers.OrderBy(static item => item, StringComparer.Ordinal).ToArray(),
            ReturnType = returnType,
            Parameters = metadata.Parameters,
            Signature = metadata.Signature,
            OverloadGroup = ownerLogicalId + ":event:" + metadata.Name,
            Invocation = new InvocationPlan
            {
                Operation = "event",
                OperationKind = ProjectionInvocationOperation.Event,
                DispatchName = metadata.Name,
                ResultType = returnType
            },
            SupportVersions = support,
            RuntimeRequirements = RuntimeRequirements("member", "event", attributes, policy),
            DocsBindingKey = DocsKey(policy, ownerLogicalId + "/contract-event/" + metadata.Name),
            Attributes = attributes,
            UseCustomEventAccessors = HasCustomEventAccessors(metadata),
            EventBackingField = ContractEventBackingField(metadata),
            Source = metadata.Source,
            Line = metadata.Line
        };
    }

    private static ProjectedType ProjectContractProjectInfo(ContractType metadata, DataProject project, IReadOnlyDictionary<string, DataLibrary> libraryById, IReadOnlyDictionary<string, DataProject> projectById, ProjectionPolicy policy)
    {
        var attributes = metadata.Attributes.Select(NormalizeAttribute).Distinct(StringComparer.Ordinal).OrderBy(static attribute => attribute, StringComparer.Ordinal).ToArray();
        var componentGuids = project.LibraryIds
            .Where(libraryById.ContainsKey)
            .Select(id => libraryById[id].Guid)
            .Where(static guid => !string.IsNullOrWhiteSpace(guid))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(1)
            .ToArray();
        var dependencies = project.ReferenceProjectIds
            .Where(projectById.ContainsKey)
            .Select(id => projectById[id])
            .Where(static dependency => dependency.Namespace.StartsWith("NetOffice.", StringComparison.Ordinal))
            .Select(static dependency => dependency.Name + "Api.dll")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var constructors = metadata.Members
            .Where(static member => member.Kind.Equals("constructor", StringComparison.OrdinalIgnoreCase))
            .OrderBy(static member => member.Signature, StringComparer.Ordinal)
            .Select((member, ordinal) =>
            {
                var logicalId = metadata.LogicalId + "/constructor/" + ordinal.ToString(CultureInfo.InvariantCulture);
                return new ConstructorPlan
                {
                    Kind = ConstructorKind(member),
                    LogicalId = logicalId,
                    IsContractOverlay = true,
                    Signature = member.Signature,
                    Accessibility = member.Accessibility,
                    Attributes = member.Attributes.Select(NormalizeAttribute).ToArray(),
                    Parameters = ParseContractParameters(member.Parameters),
                    BaseCall = ConstructorBaseCall(member.Signature),
                    RuntimeRequirements = Array.Empty<string>(),
                    DocsBindingKey = DocsKey(policy, logicalId),
                    Source = member.Source,
                    Line = member.Line
                };
            }).ToArray();
        return new ProjectedType
        {
            LogicalId = metadata.LogicalId,
            CanonicalLogicalId = metadata.LogicalId,
            ContractPartLogicalId = metadata.LogicalId,
            ContractPartSource = metadata.Source,
            DataLogicalId = "",
            Product = project.Name,
            Namespace = metadata.Namespace,
            Name = metadata.Name,
            CSharpName = metadata.Name,
            Kind = metadata.Kind,
            EntityKind = ProjectionEntityKind.Utility,
            FileCategory = ProjectionFileCategory.Utils,
            Accessibility = NormalizeTopLevelAccessibility(metadata.Accessibility),
            Modifiers = metadata.Modifiers,
            Attributes = attributes,
            BaseType = metadata.BaseType,
            Interfaces = metadata.Interfaces,
            Signature = metadata.Signature,
            SupportVersions = ParseSupport(metadata.Attributes),
            Capabilities = new[] { "project-info" },
            RuntimeRequirements = new[] { "project-info" },
            Constructors = constructors,
            ExactContractMembers = ProjectExactContractMembers(metadata, policy),
            ProjectInfo = new ProjectInfoPlan
            {
                AssemblyNamespace = project.Namespace,
                ComponentGuids = componentGuids,
                Dependencies = dependencies
            },
            DocsBindingKey = DocsKey(policy, metadata.LogicalId),
            Source = metadata.Source,
            Line = metadata.Line
        };
    }

    private static ProjectedType ProjectContractModule(ContractType metadata, string product, IReadOnlyDictionary<string, DataType[]> typeByName, ProjectionPolicy policy)
    {
        var support = ParseSupport(metadata.Attributes);
        var attributes = metadata.Attributes.Select(NormalizeAttribute).Distinct(StringComparer.Ordinal).OrderBy(static attribute => attribute, StringComparer.Ordinal).ToArray();
        var members = metadata.Members
            .Where(static member => member.Name is not "Factory" and not "Instance" and not "Invoker")
            .OrderBy(static member => member.Name, StringComparer.Ordinal).ThenBy(static member => member.Signature, StringComparer.Ordinal).ThenBy(static member => member.Source, StringComparer.Ordinal).ThenBy(static member => member.Line)
            .Select((member, ordinal) => ProjectContractModuleMember(member, metadata.LogicalId, ordinal, typeByName, policy)).ToArray();
        return new ProjectedType
        {
            LogicalId = metadata.LogicalId,
            CanonicalLogicalId = metadata.LogicalId,
            ContractPartLogicalId = metadata.LogicalId,
            ContractPartSource = metadata.Source,
            DataLogicalId = "",
            Product = product,
            Namespace = metadata.Namespace,
            Name = metadata.Name,
            CSharpName = metadata.Name,
            Kind = metadata.Kind,
            EntityKind = ProjectionEntityKind.Module,
            FileCategory = ProjectionFileCategory.Modules,
            Accessibility = NormalizeTopLevelAccessibility(metadata.Accessibility),
            Modifiers = metadata.Modifiers.Concat(new[] { "static" }).Distinct(StringComparer.Ordinal).OrderBy(static modifier => modifier, StringComparer.Ordinal).ToArray(),
            Attributes = attributes,
            Signature = metadata.Signature,
            SupportVersions = support,
            RuntimeRequirements = RuntimeRequirements("type", "module", attributes, policy),
            ExactContractMembers = ProjectExactContractMembers(metadata, policy, members),
            DocsBindingKey = DocsKey(policy, metadata.LogicalId),
            Source = metadata.Source,
            Line = metadata.Line,
            Members = members
        };
    }
    private static bool IsCollectionInfrastructureMember(ContractMember member)
        => member.Name is "GetComObjectEnumerator" or "FetchVariantComObjectEnumerator" or "GetEnumerator"
           || member.Signature.Contains("IEnumerableProvider<", StringComparison.Ordinal)
           || member.Signature.Contains("global::System.Collections.IEnumerable.GetEnumerator", StringComparison.Ordinal)
           || member.Signature.Contains("System.Collections.IEnumerable.GetEnumerator", StringComparison.Ordinal);


    private static bool IsRuntimeInfrastructureMember(ContractMember member)
        => member.Name is "InstanceType" or "LateBindingApiWrapperType"
            or "CreateEventBridge" or "DisposeEventBridge" or "EventBridgeInitialized"
            or "HasEventRecipients" or "GetEventRecipients" or "GetCountOfEventRecipients" or "RaiseCustomEvent";


    private static ProjectedMember ProjectContractModuleMember(ContractMember metadata, string ownerLogicalId, int ordinal, IReadOnlyDictionary<string, DataType[]> typeByName, ProjectionPolicy policy, bool isStatic = true)
    {
        IReadOnlyList<ProjectedParameter> parameters = ParseContractParameters(metadata.Parameters);
        var eventSinkOnly = ownerLogicalId.EndsWith("_SinkHelper", StringComparison.Ordinal);
        if (eventSinkOnly)
        {
            parameters = ApplyContractSinkArguments(parameters.ToArray(), metadata.Attributes);
            parameters = ApplyContractSinkInvocations(parameters.ToArray(), metadata.InvocationText);
        }
        var returnType = QualifyBclType(metadata.ReturnType ?? "void");
        var attributes = metadata.Attributes.Select(NormalizeAttribute).Distinct(StringComparer.Ordinal).OrderBy(static attribute => attribute, StringComparer.Ordinal).ToArray();
        var kind = metadata.Kind.ToLowerInvariant();
        var constantExpression = kind is "enumvalue" or "enum-value" or "field" or "constant" or "const" ? ContractConstantExpression(metadata.Signature) : null;
        var hasGet = kind is "property" or "indexer" && metadata.InvocationText.Any(static text => text.Contains("PropertyGet", StringComparison.Ordinal));
        var hasSet = kind is "property" or "indexer" && metadata.InvocationText.Any(static text => text.Contains("PropertySet", StringComparison.Ordinal));
        var logicalId = ownerLogicalId + "/contract/" + metadata.Name + "/" + ordinal.ToString(CultureInfo.InvariantCulture);
        var fallbackOperation = kind is "property" or "indexer" ? ProjectionInvocationOperation.PropertyGet : kind == "event" ? ProjectionInvocationOperation.Event : ProjectionInvocationOperation.Method;
        var primaryOperation = hasGet
            ? ProjectionInvocationOperation.PropertyGet
            : hasSet ? ContractInvocationOperation(metadata.InvocationText, ProjectionInvocationOperation.PropertySet)
            : ContractInvocationOperation(metadata.InvocationText, fallbackOperation);
        var getterText = ContractOperationText(metadata.InvocationText, primaryOperation);
        var setterText = ContractOperationText(metadata.InvocationText, ProjectionInvocationOperation.PropertySet);
        var conversion = ContractReturnConversion(string.Join("\n", getterText), metadata, returnType, typeByName);
        var argumentOrder = ContractArgumentOrder(getterText, parameters);
        var arguments = OrderedInvocationArguments(parameters, argumentOrder);
        arguments = OverlayContractArgumentExpressions(arguments, getterText);
        var getterLocalReturn = ContractLocalReturn(getterText);
        var setterLocalReturn = ContractLocalReturn(setterText);
        if (primaryOperation is ProjectionInvocationOperation.PropertySet or ProjectionInvocationOperation.PropertySetVariant or ProjectionInvocationOperation.PropertySetEnum or ProjectionInvocationOperation.PropertyPutRef)
            arguments = arguments.Select(static argument => argument.Expression.Equals("value", StringComparison.Ordinal) ? argument with { IsPropertyValue = true } : argument).ToArray();
        var getPlan = new InvocationPlan
        {
            Operation = kind is "property" or "indexer" ? "property-get" : kind == "event" ? "event" : "method",
            OperationKind = primaryOperation,
            Api = ContractInvocationApi(getterText, arguments.Any(static argument => argument.ByRef) ? ProjectionInvocationApi.Invoker : ProjectionInvocationApi.Factory),
            Target = ContractInvocationTarget(getterText) ?? (isStatic ? "_instance" : "this"),
            DispatchName = ContractDispatchName(getterText) ?? metadata.Name,
            ArgumentCount = argumentOrder.Count,
            ArgumentOrder = argumentOrder,
            ResultType = returnType,
            ReturnConversion = conversion,
            FactoryMethodSuffix = ContractFactoryMethodSuffix(getterText) ?? ScalarFactorySuffix(returnType),
            ArgumentPacking = ContractArgumentPacking(getterText),
            ObjectArrayStyle = ContractObjectArrayStyle(getterText),
            RawCallCast = ContractRawCallCast(getterText),
            KnownReferenceFactoryStyle = ContractKnownReferenceFactoryStyle(getterText),
            ReturnValueStyle = getterLocalReturn.Name is null ? ProjectionReturnValueStyle.Direct : ProjectionReturnValueStyle.Local,
            LocalReturnName = getterLocalReturn.Name,
            LocalReturnType = getterLocalReturn.Type,
            ReturnCastStyle = ContractReturnCastStyle(getterText),
            CallKind = ContractInvocationCallKind(getterText),
            InvokerCallStyle = ContractInvokerCallStyle(getterText),
            RequiresProxy = conversion is ProjectionReturnConversion.KnownReference or ProjectionReturnConversion.Reference or ProjectionReturnConversion.UntypedReference or ProjectionReturnConversion.BaseReference,
            HasContractInvocation = getterText.Count != 0,
            Arguments = arguments,
            ReleaseArguments = ContractReleaseArguments(getterText)
        };
        var redirect = metadata.Attributes.Select(static attribute => RedirectRegex.Match(attribute)).FirstOrDefault(static match => match.Success);
        if (redirect is not null)
        {
            var target = redirect.Groups["target"].Value;
            getPlan = getPlan with
            {
                Operation = "local-forward",
                OperationKind = ProjectionInvocationOperation.LocalForward,
                Api = ProjectionInvocationApi.Local,
                Target = null,
                DispatchName = target,
                ReturnConversion = ProjectionReturnConversion.None,
                FactoryMethodSuffix = null,
                RequiresProxy = false,
                HasContractInvocation = false
            };
        }
        else if (metadata.Modifiers.Contains("static", StringComparer.Ordinal)
                 && metadata.Name is "GetActiveInstance" or "GetActiveInstances")
        {
            getPlan = getPlan with
            {
                Operation = metadata.Name == "GetActiveInstance" ? "active-instance" : "active-instances",
                OperationKind = metadata.Name == "GetActiveInstance" ? ProjectionInvocationOperation.ActiveInstance : ProjectionInvocationOperation.ActiveInstances,
                Api = ProjectionInvocationApi.Runtime,
                Target = null,
                ReturnConversion = ProjectionReturnConversion.None,
                FactoryMethodSuffix = null,
                RequiresProxy = false,
                HasContractInvocation = false
            };
        }
        var accessorPlans = new List<InvocationPlan>();
        if (hasGet) accessorPlans.Add(getPlan);
        if (hasSet) accessorPlans.Add(getPlan with
        {
            Operation = "property-set",
            OperationKind = ProjectionInvocationOperation.PropertySet,
            Api = ContractInvocationApi(setterText, ProjectionInvocationApi.Factory),
            ArgumentCount = ContractArgumentOrder(setterText, parameters, true).Count,
            ArgumentOrder = ContractArgumentOrder(setterText, parameters, true),
            ResultType = "void",
            ReturnConversion = ProjectionReturnConversion.None,
            SetValueConversion = ContractReturnConversion(string.Join("\n", setterText), metadata, returnType, typeByName),
            FactoryMethodSuffix = ContractFactoryMethodSuffix(setterText),
            ArgumentPacking = ContractArgumentPacking(setterText),
            ObjectArrayStyle = ContractObjectArrayStyle(setterText),
            CallKind = ContractInvocationCallKind(setterText),
            RawCallCast = ContractRawCallCast(setterText),
            KnownReferenceFactoryStyle = ContractKnownReferenceFactoryStyle(setterText),
            ReturnValueStyle = setterLocalReturn.Name is null ? ProjectionReturnValueStyle.Direct : ProjectionReturnValueStyle.Local,
            LocalReturnName = setterLocalReturn.Name,
            LocalReturnType = setterLocalReturn.Type,
            ReturnCastStyle = ContractReturnCastStyle(setterText),
            InvokerCallStyle = ContractInvokerCallStyle(setterText),
            RequiresProxy = false,
            HasContractInvocation = setterText.Count != 0,
            Arguments = OverlayContractArgumentExpressions(
                OrderedInvocationArguments(parameters, ContractArgumentOrder(setterText, parameters, true), true, returnType),
                setterText),
            ReleaseArguments = ContractReleaseArguments(setterText)
        });
        return new ProjectedMember
        {
            LogicalId = logicalId,
            DataLogicalId = logicalId,
            EmissionOwnerLogicalId = ownerLogicalId,
            IsContractDerived = true,
            Name = metadata.Name,
            CSharpName = Sanitize(metadata.Name, policy.Names),
            Kind = kind,
            Accessibility = metadata.Accessibility,
            Modifiers = (isStatic ? metadata.Modifiers.Concat(new[] { "static" }) : metadata.Modifiers).Distinct(StringComparer.Ordinal).OrderBy(static modifier => modifier, StringComparer.Ordinal).ToArray(),
            ReturnType = returnType,
            Parameters = metadata.Parameters,
            ParameterList = parameters,
            Signature = metadata.Signature,
            OverloadGroup = ownerLogicalId + ":overload:" + metadata.Name,
            RuntimeMemberKind = ContractRuntimeMemberKind(metadata),
            Invocation = getPlan,
            AccessorInvocations = accessorPlans,
            Accessors = new[] { hasGet ? "get" : null, hasSet ? "set" : null }.Where(static accessor => accessor is not null).Cast<string>().ToArray(),
            SupportVersions = ParseSupport(metadata.Attributes),
            RuntimeRequirements = RuntimeRequirements("member", kind, attributes, policy),
            DocsBindingKey = DocsKey(policy, logicalId),
            Attributes = attributes,
            ConstantValue = constantExpression,
            ConstantExpression = constantExpression,
            ValueType = metadata.ReturnType,
            EventSinkOnly = eventSinkOnly,
            EventValidationKey = eventSinkOnly ? ContractEventValidationKey(metadata.InvocationText) : null,
            EventInvalidReleaseArguments = eventSinkOnly ? ContractEventInvalidReleaseArguments(metadata.InvocationText, parameters) : null,
            EventValidationInlineReturn = eventSinkOnly && ContractEventValidationInlineReturn(metadata.InvocationText),
            Source = metadata.Source,
            Line = metadata.Line
        };
    }

    private static IReadOnlyList<ProjectedParameter> ParseContractParameters(string? parameters)
    {
        if (string.IsNullOrWhiteSpace(parameters) || parameters.Trim() is "()" or "[]") return Array.Empty<ProjectedParameter>();
        var value = parameters.Trim();
        if ((value.StartsWith("(", StringComparison.Ordinal) && value.EndsWith(")", StringComparison.Ordinal))
            || (value.StartsWith("[", StringComparison.Ordinal) && value.EndsWith("]", StringComparison.Ordinal))) value = value[1..^1];
        return SplitTopLevel(value).Select((text, index) => ParseParameterDeclaration(text, index, true)).ToArray();
    }

    private static string? ContractConstantExpression(string signature)
    {
        var equals = signature.IndexOf('=', StringComparison.Ordinal);
        if (equals < 0) return null;
        var expression = signature[(equals + 1)..].Trim().TrimEnd(';', ',').Trim();
        return expression.Length == 0 ? null : expression;
    }

    private static ProjectedMember ProjectValue(DataValue value, TypeDescriptor descriptor, IReadOnlyList<VersionSupport> support, ProjectionPolicy policy, string ownerLogicalId, ContractMember? metadata)
    {
        var isEnum = descriptor.EntityKind == ProjectionEntityKind.Enum;
        var kind = metadata?.Kind ?? (isEnum ? "enum-value" : "field");
        var modifiers = metadata?.Modifiers ?? (isEnum ? Array.Empty<string>() : new[] { "const", "static" });
        var expression = metadata is null ? ConstantExpression(value.Value, value.ValueType) : ContractConstantExpression(metadata.Signature) ?? ConstantExpression(value.Value, value.ValueType);
        var name = metadata?.Name ?? value.Name;
        var returnType = metadata?.ReturnType ?? value.ValueType;
        return new ProjectedMember
        {
            LogicalId = ownerLogicalId + "/" + value.LogicalId,
            DataLogicalId = value.LogicalId,
            EmissionOwnerLogicalId = ownerLogicalId,
            Name = name,
            CSharpName = Sanitize(name, policy.Names),
            Kind = kind,
            Accessibility = metadata?.Accessibility ?? "public",
            Modifiers = modifiers,
            ReturnType = returnType,
            ValueType = returnType,
            ConstantValue = value.Value,
            ConstantExpression = expression,
            Signature = metadata?.Signature ?? (isEnum ? $"{Sanitize(value.Name, policy.Names)} = {expression}" : $"public const {value.ValueType ?? "object"} {Sanitize(value.Name, policy.Names)} = {expression}"),
            OverloadGroup = ownerLogicalId + ":value:" + Sanitize(value.Name, policy.Names),
            Invocation = new InvocationPlan { Operation = "field", OperationKind = ProjectionInvocationOperation.Field, DispatchName = value.Name, ReturnConversion = ReturnConversion(returnType, null, EmptyTypes) },
            SupportVersions = support,
            DocsBindingKey = DocsKey(policy, value.LogicalId),
            Attributes = metadata is null ? ContractOrSupportAttributes(null, support) : ContractOrSupportAttributes(metadata.Attributes, support),
            Source = metadata?.Source ?? DataLocation(value.Provenance),
            Line = metadata?.Line ?? SourceLine(value.Provenance)
        };
    }

    private static string ConstantExpression(string value, string? type)
    {
        if (string.Equals(type, "string", StringComparison.OrdinalIgnoreCase)) return System.Text.Json.JsonSerializer.Serialize(value);
        if (string.Equals(type, "char", StringComparison.OrdinalIgnoreCase))
        {
            var escaped = System.Text.Json.JsonSerializer.Serialize(value);
            return value.Length == 1 ? "'" + escaped[1..^1].Replace("'", "\\'", StringComparison.Ordinal) + "'" : escaped;
        }
        if (string.Equals(type, "bool", StringComparison.OrdinalIgnoreCase) || string.Equals(type, "boolean", StringComparison.OrdinalIgnoreCase))
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "-1" ? "true" : "false";
        return value;
    }

    private static InvocationPlan Invocation(DataMember data, InvocationEvidence? evidence, ContractMember? metadata, IReadOnlyList<ProjectedParameter> parameters, string? returnType, string projectedKind, IReadOnlyDictionary<string, DataType[]> typeByName, IReadOnlyDictionary<string, string> qualifiedTypeById)
    {
        var operationKind = projectedKind == "method" && data.AccessorKind is not null
            ? ProjectionInvocationOperation.Method
            : data.AccessorKind switch
            {
                "get" => ProjectionInvocationOperation.PropertyGet,
                "put" => ProjectionInvocationOperation.PropertySet,
                "putref" => ProjectionInvocationOperation.PropertyPutRef,
                _ when data.Kind.Contains("event", StringComparison.OrdinalIgnoreCase) => ProjectionInvocationOperation.Event,
                _ when data.Kind.Contains("property", StringComparison.OrdinalIgnoreCase) => ProjectionInvocationOperation.PropertyGet,
                _ when data.Kind.Contains("field", StringComparison.OrdinalIgnoreCase) => ProjectionInvocationOperation.Field,
                _ => ProjectionInvocationOperation.Method
            };
        var operation = operationKind switch
        {
            ProjectionInvocationOperation.PropertyGet => "property-get",
            ProjectionInvocationOperation.PropertySet => "property-set",
            ProjectionInvocationOperation.PropertyPutRef => "property-putref",
            ProjectionInvocationOperation.Event => "event",
            ProjectionInvocationOperation.Field => "field",
            _ => evidence?.Operation?.ToLowerInvariant() == "property" && data.AccessorKind is null ? "property-get" : "method"
        };
        var isSetter = operationKind is ProjectionInvocationOperation.PropertySet or ProjectionInvocationOperation.PropertyPutRef;
        var argumentOrder = parameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName).ToList();
        if (isSetter) argumentOrder.Add("value");
        var resultType = isSetter ? "void" : QualifiedReferenceType(data.ReturnTypeReference, evidence?.ResultType ?? returnType, qualifiedTypeById);
        var resultReference = isSetter ? null : data.ReturnTypeReference;
        var invocationArguments = InvocationArguments(parameters, isSetter, returnType);
        return new InvocationPlan
        {
            Operation = operation,
            OperationKind = operationKind,
            Api = invocationArguments.Any(static argument => argument.ByRef) ? ProjectionInvocationApi.Invoker : ProjectionInvocationApi.Factory,
            DispatchName = evidence?.DispatchName ?? metadata?.Name ?? data.Name,
            DispId = evidence?.DispId ?? data.DispId,
            ArgumentCount = argumentOrder.Count,
            ArgumentOrder = argumentOrder,
            ResultType = resultType,
            ResultTypeReference = ProjectReference(resultReference, qualifiedTypeById),
            SetValueConversion = isSetter ? ReturnConversion(data.ReturnType, data.ReturnTypeReference, typeByName) : null,
            ReturnConversion = ReturnConversion(resultType, resultReference, typeByName),
            FactoryMethodSuffix = ScalarFactorySuffix(resultType),
            RequiresProxy = !isSetter && (data.IsComProxy || data.ReturnTypeReference?.IsComProxy == true || (evidence?.RequiresProxy ?? (ReturnConversion(returnType, data.ReturnTypeReference, typeByName) is ProjectionReturnConversion.KnownReference or ProjectionReturnConversion.Reference))),
            HasContractInvocation = metadata?.InvocationText.Count > 0,
            Arguments = invocationArguments,
            ReleaseArguments = invocationArguments.Any(static argument => argument.ByRef)
        };
    }

    private static InvocationPlan OverlayContractInvocation(InvocationPlan plan, ContractMember? metadata, IReadOnlyList<ProjectedParameter> parameters, string returnType, IReadOnlyDictionary<string, DataType[]> typeByName)
    {
        if (metadata is null) return plan;
        var redirect = metadata.Attributes.Select(static attribute => RedirectRegex.Match(attribute)).FirstOrDefault(static match => match.Success);
        if (redirect is not null)
        {
            var target = redirect.Groups["target"].Value;
            var arguments = InvocationArguments(parameters, false, returnType);
            return plan with
            {
                Operation = "local-forward",
                OperationKind = ProjectionInvocationOperation.LocalForward,
                Api = ProjectionInvocationApi.Local,
                Target = null,
                DispatchName = target,
                ArgumentCount = parameters.Count,
                ArgumentOrder = parameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName).ToArray(),
                ReturnConversion = ProjectionReturnConversion.None,
                FactoryMethodSuffix = null,
                SetValueConversion = null,
                RequiresProxy = false,
                HasContractInvocation = false,
                Arguments = arguments,
                ReleaseArguments = false
            };
        }
        var selectedText = ContractOperationText(metadata.InvocationText, plan.OperationKind);
        var text = string.Join("\n", selectedText);
        var conversion = ContractReturnConversion(text, metadata, returnType, typeByName, plan.ReturnConversion);
        var planIsSetter = plan.OperationKind is ProjectionInvocationOperation.PropertySet or ProjectionInvocationOperation.PropertySetVariant or ProjectionInvocationOperation.PropertySetEnum or ProjectionInvocationOperation.PropertyPutRef;
        var operationKind = planIsSetter && text.Contains("ExecuteVariantPropertySet", StringComparison.Ordinal) ? ProjectionInvocationOperation.PropertySetVariant
            : planIsSetter && text.Contains("ExecuteEnumPropertySet", StringComparison.Ordinal) ? ProjectionInvocationOperation.PropertySetEnum
            : planIsSetter && text.Contains("ExecuteReferencePropertySet", StringComparison.Ordinal) ? ProjectionInvocationOperation.PropertyPutRef
            : planIsSetter ? plan.OperationKind : ContractInvocationOperation(selectedText, plan.OperationKind);
        var isSetter = operationKind is ProjectionInvocationOperation.PropertySet or ProjectionInvocationOperation.PropertySetVariant or ProjectionInvocationOperation.PropertySetEnum or ProjectionInvocationOperation.PropertyPutRef;
        var argumentOrder = ContractArgumentOrder(selectedText, parameters, isSetter);
        var invocationArguments = OrderedInvocationArguments(parameters, argumentOrder, isSetter, returnType);
        var localReturn = ContractLocalReturn(selectedText);
        return plan with
        {
            OperationKind = operationKind,
            Api = ContractInvocationApi(selectedText, plan.Api),
            Target = ContractInvocationTarget(selectedText) ?? plan.Target,
            ArgumentCount = argumentOrder.Count,
            ArgumentOrder = argumentOrder,
            ReturnConversion = isSetter ? ProjectionReturnConversion.None : conversion,
            ResultTypeReference = null,
            FactoryMethodSuffix = ContractFactoryMethodSuffix(selectedText) ?? plan.FactoryMethodSuffix,
            ArgumentPacking = ContractArgumentPacking(selectedText),
            ObjectArrayStyle = ContractObjectArrayStyle(selectedText),
            SetValueConversion = isSetter ? conversion : null,
            RawCallCast = ContractRawCallCast(selectedText),
            KnownReferenceFactoryStyle = ContractKnownReferenceFactoryStyle(selectedText),
            ReturnValueStyle = localReturn.Name is null ? ProjectionReturnValueStyle.Direct : ProjectionReturnValueStyle.Local,
            LocalReturnName = localReturn.Name,
            LocalReturnType = localReturn.Type,
            ReturnCastStyle = ContractReturnCastStyle(selectedText),
            CallKind = ContractInvocationCallKind(selectedText),
            InvokerCallStyle = ContractInvokerCallStyle(selectedText),
            RequiresProxy = !isSetter && conversion is ProjectionReturnConversion.KnownReference or ProjectionReturnConversion.Reference or ProjectionReturnConversion.UntypedReference or ProjectionReturnConversion.BaseReference,
            HasContractInvocation = selectedText.Count != 0,
            Arguments = OverlayContractArgumentExpressions(invocationArguments, selectedText),
            ReleaseArguments = ContractReleaseArguments(selectedText)
        };
    }

    private static ProjectionInvocationOperation ContractInvocationOperation(IReadOnlyList<string> invocationText, ProjectionInvocationOperation fallback)
    {
        var text = string.Join("\n", invocationText);
        if (text.Contains("VariantPropertySet", StringComparison.Ordinal)) return ProjectionInvocationOperation.PropertySetVariant;
        if (text.Contains("EnumPropertySet", StringComparison.Ordinal)) return ProjectionInvocationOperation.PropertySetEnum;
        if (text.Contains("ReferencePropertySet", StringComparison.Ordinal)) return ProjectionInvocationOperation.PropertyPutRef;
        if (text.Contains("PropertySet", StringComparison.Ordinal)) return ProjectionInvocationOperation.PropertySet;
        if (text.Contains("PropertyGet", StringComparison.Ordinal)) return ProjectionInvocationOperation.PropertyGet;
        if (text.Contains("MethodGet", StringComparison.Ordinal)) return ProjectionInvocationOperation.Method;
        return fallback;
    }

    private static string? ContractDispatchName(IEnumerable<string> invocationText)
    {
        foreach (var text in invocationText)
        {
            var match = ContractInvocationRegex.Match(text);
            if (match.Success) return match.Groups["dispatch"].Value;
            match = InvokerDispatchRegex.Match(text);
            if (match.Success) return match.Groups["dispatch"].Value;
        }
        return null;
    }
    private static string? ContractInvocationTarget(IEnumerable<string> invocationText)
    {
        foreach (var text in invocationText)
        {
            var match = ContractInvocationRegex.Match(text);
            if (match.Success) return match.Groups["target"].Value;
            match = InvokerDispatchRegex.Match(text);
            if (match.Success) return match.Groups["target"].Value;
        }
        return null;
    }


    private static IReadOnlyList<string> ContractArgumentOrder(IReadOnlyList<string> invocationText, IReadOnlyList<ProjectedParameter> parameters, bool includePropertyValue = false)
    {
        foreach (var text in invocationText)
        {
            var match = ContractInvocationRegex.Match(text);
            if (!match.Success || !match.Groups["arguments"].Success) continue;
            var arguments = SplitTopLevel(match.Groups["arguments"].Value)
                .Select(static argument => argument.Trim())
                .Where(static argument => !argument.EndsWith(".LateBindingApiWrapperType", StringComparison.Ordinal))
                .ToArray();
            if (arguments.Length == 1)
            {
                var arrayMatch = Regex.Match(arguments[0], @"^new\s+object\s*\[\]\s*\{(?<items>.*)\}$", RegexOptions.CultureInvariant);
                if (arrayMatch.Success)
                    arguments = SplitTopLevel(arrayMatch.Groups["items"].Value)
                        .Select(static argument => argument.Trim())
                        .ToArray();
            }
            if (arguments.Length != 0) return arguments;
        }
        var fallback = parameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName);
        return includePropertyValue ? fallback.Append("value").ToArray() : fallback.ToArray();
    }

    private static IReadOnlyList<InvocationArgumentPlan> OrderedInvocationArguments(IReadOnlyList<ProjectedParameter> parameters, IReadOnlyList<string> argumentOrder, bool isSetter = false, string? setterType = null)
    {
        var arguments = InvocationArguments(parameters, isSetter, setterType).ToDictionary(static argument => argument.Expression, StringComparer.Ordinal);
        return argumentOrder.Select(argument => arguments.TryGetValue(argument, out var known)
            ? known
            : new InvocationArgumentPlan { Expression = argument }).ToArray();
    }

    private static IReadOnlyList<InvocationArgumentPlan> OverlayContractArgumentExpressions(IReadOnlyList<InvocationArgumentPlan> arguments, IReadOnlyList<string> invocationText)
    {
        foreach (var line in invocationText)
        {
            var match = ValidateParamsRegex.Match(line);
            if (!match.Success) continue;
            var expressions = SplitTopLevel(match.Groups["arguments"].Value).Select(static value => value.Trim()).ToArray();
            if (expressions.Length != arguments.Count) continue;
            return arguments.Select((argument, index) => argument with
            {
                Expression = expressions[index],
                InitializationExpression = ContractInitializationExpression(argument.Expression, invocationText)
            }).ToArray();
        }
        return arguments;
    }

    private static string? ContractInitializationExpression(string parameterName, IReadOnlyList<string> invocationText)
    {
        foreach (var line in invocationText)
        {
            var match = Regex.Match(line, @"^\s*" + Regex.Escape(parameterName) + @"\s*=\s*(?<expression>.+?)\s*;\s*$", RegexOptions.CultureInvariant);
            if (match.Success) return match.Groups["expression"].Value;
        }
        return null;
    }

    private static bool ContractReleaseArguments(IReadOnlyList<string> invocationText)
        => invocationText.Any(static line => ReleaseParamsRegex.IsMatch(line));

    private static ProjectionReturnCastStyle ContractReturnCastStyle(IEnumerable<string> invocationText)
    {
        var text = string.Join("\n", invocationText);
        if (Regex.IsMatch(text, @"return\s+\([A-Za-z0-9_.:<>\[\]]+\)\s*returnItem", RegexOptions.CultureInvariant)
            || Regex.IsMatch(text, @"=\s*\([A-Za-z0-9_.:<>\[\]]+\)\s*Factory\.Create", RegexOptions.CultureInvariant))
            return ProjectionReturnCastStyle.Explicit;
        if (Regex.IsMatch(text, @"(?:return\s+returnItem|Factory\.Create[A-Za-z0-9]+FromComProxy\([^;]+\))\s+as\s+[A-Za-z0-9_.:<>\[\]]+", RegexOptions.CultureInvariant))
            return ProjectionReturnCastStyle.As;
        return ProjectionReturnCastStyle.Default;
    }

    private static ProjectionRawCallCast ContractRawCallCast(IEnumerable<string> invocationText)
        => invocationText.Any(static line => Regex.IsMatch(line, @"\breturnItem\s*=\s*\(object\)\s*Invoker\.", RegexOptions.CultureInvariant))
            ? ProjectionRawCallCast.Object
            : ProjectionRawCallCast.None;

    private static ProjectionKnownReferenceFactoryStyle ContractKnownReferenceFactoryStyle(IEnumerable<string> invocationText)
        => invocationText.Any(static line => line.Contains("CreateKnownObjectFromComProxy(", StringComparison.Ordinal))
            ? ProjectionKnownReferenceFactoryStyle.NonGeneric
            : ProjectionKnownReferenceFactoryStyle.Generic;

    private static (string? Name, string? Type) ContractLocalReturn(IEnumerable<string> invocationText)
    {
        foreach (var line in invocationText)
        {
            var match = SinkLocalDeclarationRegex.Match(line);
            if (!match.Success) continue;
            var expression = match.Groups["expression"].Value;
            if (expression.Contains("CreateKnownObjectFromComProxy", StringComparison.Ordinal)
                || expression.Contains("CreateObjectFromComProxy", StringComparison.Ordinal)
                || expression.Contains("CreateEventArgumentObjectFromComProxy", StringComparison.Ordinal))
                return (match.Groups["local"].Value, match.Groups["type"].Value);
        }
        return (null, null);
    }


    private static string? ContractEventValidationKey(IReadOnlyList<string>? invocationText)
    {
        if (invocationText is null) return null;
        foreach (var line in invocationText)
        {
            var match = EventValidationRegex.Match(line);
            if (match.Success) return match.Groups["key"].Value;
        }
        return null;
    }

    private static IReadOnlyList<string>? ContractEventInvalidReleaseArguments(IReadOnlyList<string>? invocationText, IReadOnlyList<ProjectedParameter> parameters)
    {
        if (invocationText is null) return null;
        var knownNames = parameters.Select(static parameter => parameter.CSharpName)
            .Concat(parameters.Select(static parameter => parameter.Name))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var line in invocationText)
        {
            var match = ReleaseParamsRegex.Match(line);
            if (!match.Success) continue;
            var value = match.Groups["arguments"].Value.Trim();
            if (value.Length == 0) return Array.Empty<string>();
            var arguments = SplitTopLevel(value).Select(static argument => argument.Trim()).ToArray();
            return arguments.All(knownNames.Contains) ? arguments : null;
        }
        return null;
    }

    private static bool ContractEventValidationInlineReturn(IReadOnlyList<string>? invocationText)
        => invocationText?.Any(static line => ReleaseParamsRegex.IsMatch(line)
                                              && line.Contains("return;", StringComparison.Ordinal)) == true;

    private static ProjectionInvokerCallStyle ContractInvokerCallStyle(IEnumerable<string> invocationText)
        => invocationText.Any(static line => Regex.IsMatch(line, @"^\s*return\s+Invoker\.(?:PropertyGet|MethodReturn)\s*\(", RegexOptions.CultureInvariant)
                                                  && !line.Contains("paramsArray", StringComparison.Ordinal))
            ? ProjectionInvokerCallStyle.Direct
            : ProjectionInvokerCallStyle.ParamsArray;

    private static IReadOnlyList<string> ContractOperationText(IReadOnlyList<string> invocationText, ProjectionInvocationOperation operation)
    {
        var hasGetter = invocationText.Any(static text => text.Contains("PropertyGet", StringComparison.Ordinal)
                                                          || text.Contains("MethodGet", StringComparison.Ordinal)
                                                          || text.Contains("MethodReturn", StringComparison.Ordinal));
        var hasSetter = invocationText.Any(static text => text.Contains("PropertySet", StringComparison.Ordinal));
        if (!(hasGetter && hasSetter)) return invocationText;
        bool Matches(string text) => operation switch
        {
            ProjectionInvocationOperation.PropertyGet => text.Contains("PropertyGet", StringComparison.Ordinal),
            ProjectionInvocationOperation.PropertySet or ProjectionInvocationOperation.PropertySetVariant or ProjectionInvocationOperation.PropertySetEnum or ProjectionInvocationOperation.PropertyPutRef => text.Contains("PropertySet", StringComparison.Ordinal),
            ProjectionInvocationOperation.Method => text.Contains("MethodGet", StringComparison.Ordinal) || text.Contains("MethodReturn", StringComparison.Ordinal),
            _ => true
        };
        var selected = invocationText.Where(Matches).ToArray();
        return selected.Length == 0 ? invocationText : selected;
    }

    private static ProjectionReturnConversion ContractReturnConversion(string text, ContractMember metadata, string returnType, IReadOnlyDictionary<string, DataType[]> typeByName, ProjectionReturnConversion? fallback = null)
        => text.Contains("ExecuteKnownReference", StringComparison.Ordinal) ? ProjectionReturnConversion.KnownReference
            : text.Contains("ExecuteBaseReference", StringComparison.Ordinal) ? ProjectionReturnConversion.BaseReference
            : text.Contains("ExecuteReference", StringComparison.Ordinal)
                ? text.Contains("ExecuteReferenceMethodGet<", StringComparison.Ordinal) || text.Contains("ExecuteReferencePropertyGet<", StringComparison.Ordinal)
                    ? ProjectionReturnConversion.Reference
                    : ProjectionReturnConversion.UntypedReference
            : text.Contains("ExecuteEnum", StringComparison.Ordinal) ? ProjectionReturnConversion.Enum
            : text.Contains("ExecuteString", StringComparison.Ordinal) ? ProjectionReturnConversion.String
            : text.Contains("ExecuteArray", StringComparison.Ordinal) ? ProjectionReturnConversion.Array
            : text.Contains("ExecuteStruct", StringComparison.Ordinal) ? ProjectionReturnConversion.Struct
            : text.Contains("ExecuteVariant", StringComparison.Ordinal) ? ProjectionReturnConversion.Variant
            : text.Contains("CreateObjectFromComProxy", StringComparison.Ordinal) ? ProjectionReturnConversion.UntypedReference
            : text.Contains("ExecuteObject", StringComparison.Ordinal) ? ProjectionReturnConversion.Native
            : text.Contains("ExecuteValue", StringComparison.Ordinal) ? ProjectionReturnConversion.Value
            : text.Contains("ExecuteBool", StringComparison.Ordinal)
              || text.Contains("ExecuteByte", StringComparison.Ordinal)
              || text.Contains("ExecuteInt", StringComparison.Ordinal)
              || text.Contains("ExecuteSingle", StringComparison.Ordinal)
              || text.Contains("ExecuteDouble", StringComparison.Ordinal)
              || text.Contains("ExecuteDecimal", StringComparison.Ordinal)
              || text.Contains("ExecuteDateTime", StringComparison.Ordinal)
                ? ProjectionReturnConversion.Scalar
                : metadata.Attributes.Any(static attribute => attribute.Contains("NativeResult", StringComparison.Ordinal))
                  || returnType.StartsWith("stdole.", StringComparison.OrdinalIgnoreCase)
                    ? ProjectionReturnConversion.Native
                    : fallback ?? ReturnConversion(returnType, null, typeByName);

    private static string? ContractFactoryMethodSuffix(IEnumerable<string> invocationText)
    {
        foreach (var text in invocationText)
        {
            var match = Regex.Match(text, @"(?:Factory|Invoker)\.Execute(?<suffix>[A-Za-z0-9]+?)(?:MethodGet|PropertyGet|PropertySet)\b", RegexOptions.CultureInvariant);
            if (match.Success) return match.Groups["suffix"].Value;
        }
        return null;
    }

    private static IReadOnlyList<InvocationArgumentPlan> InvocationArguments(IReadOnlyList<ProjectedParameter> parameters, bool isSetter, string? setterType)
    {
        var result = parameters.OrderBy(static parameter => parameter.Position).Select(parameter => new InvocationArgumentPlan
        {
            Expression = parameter.CSharpName,
            ByRef = parameter.RefKind is "ref" or "out",
            InitializationExpression = parameter.RefKind == "out" ? "default(" + parameter.Type + ")" : null,
            WriteBackExpression = parameter.RefKind is "ref" or "out" ? parameter.CSharpName : null,
            WriteBackType = parameter.Type,
            WriteBackConversion = parameter.SinkArgument?.ConversionExpression
        }).ToList();
        if (isSetter)
            result.Add(new InvocationArgumentPlan { Expression = "value", WriteBackType = setterType ?? "object", IsPropertyValue = true });
        return result;
    }

    private static ProjectionReturnConversion ReturnConversion(string? type, DataTypeReference? reference, IReadOnlyDictionary<string, DataType[]> typeByName)
    {
        if (reference?.IsArray == true) return ProjectionReturnConversion.Array;
        if (reference?.IsEnum == true) return ProjectionReturnConversion.Enum;
        if (reference?.IsNative == true && reference.TargetTypeId is not null) return ProjectionReturnConversion.Native;
        if (reference?.IsComProxy == true && !string.IsNullOrWhiteSpace(reference.TypeKey)) return ProjectionReturnConversion.KnownReference;
        if (reference?.IsComProxy == true) return ProjectionReturnConversion.Reference;
        if (reference?.TypeKind?.Contains("record", StringComparison.OrdinalIgnoreCase) == true || reference?.TypeKind?.Contains("alias", StringComparison.OrdinalIgnoreCase) == true) return ProjectionReturnConversion.Value;
        if (string.IsNullOrWhiteSpace(type) || string.Equals(type, "void", StringComparison.OrdinalIgnoreCase)) return ProjectionReturnConversion.None;
        var value = type.Trim().TrimEnd('?');
        if (value.EndsWith("[]", StringComparison.Ordinal)) return ProjectionReturnConversion.Array;
        if (string.Equals(value, "string", StringComparison.OrdinalIgnoreCase)) return ProjectionReturnConversion.String;
        if (ScalarTypes.Contains(value)) return ProjectionReturnConversion.Scalar;
        var shortName = value.Split('.').Last();
        if (typeByName.TryGetValue(shortName, out var candidates))
        {
            if (candidates.Any(static item => item.Kind.Equals("enum", StringComparison.OrdinalIgnoreCase))) return ProjectionReturnConversion.Enum;
            if (candidates.Any(static item => item.Kind is "Record" or "Struct" or "Alias")) return ProjectionReturnConversion.Value;
            return ProjectionReturnConversion.KnownReference;
        }
        return string.Equals(value, "object", StringComparison.OrdinalIgnoreCase) ? ProjectionReturnConversion.UntypedReference : ProjectionReturnConversion.Value;
    }

    private static string? ScalarFactorySuffix(string? type)
    {
        if (string.IsNullOrWhiteSpace(type)) return null;
        return TypeIdentity(type).ToLowerInvariant() switch
        {
            "bool" or "boolean" => "Bool",
            "byte" => "Byte",
            "sbyte" => "SByte",
            "short" or "int16" => "Int16",
            "ushort" or "uint16" => "UInt16",
            "int" or "int32" => "Int32",
            "uint" or "uint32" => "UInt32",
            "long" or "int64" => "Int64",
            "ulong" or "uint64" => "UInt64",
            "float" or "single" => "Single",
            "double" => "Double",
            "decimal" => "Decimal",
            "char" => "Char",
            "string" => "String",
            "datetime" => "DateTime",
            _ => null
        };
    }

    private static ProjectedTypeReference? ProjectReference(DataTypeReference? reference, IReadOnlyDictionary<string, string> qualifiedTypeById)
        => reference is null ? null : new ProjectedTypeReference
        {
            Name = reference.Name,
            TypeKind = reference.TypeKind,
            TargetTypeId = reference.TargetTypeId,
            QualifiedName = QualifiedReferenceType(reference, reference.Name, qualifiedTypeById),
            VarType = reference.VarType,
            MarshalAs = reference.MarshalAs,
            TypeKey = reference.TypeKey,
            ProjectKey = reference.ProjectKey,
            LibraryKey = reference.LibraryKey,
            IsComProxy = reference.IsComProxy,
            IsExternal = reference.IsExternal,
            IsEnum = reference.IsEnum,
            IsArray = reference.IsArray,
            IsNative = reference.IsNative
        };

    private static string QualifiedReferenceType(DataTypeReference? reference, string? fallback, IReadOnlyDictionary<string, string> qualifiedTypeById)
    {
        if (reference is not null && reference.Name.Equals("Guid", StringComparison.OrdinalIgnoreCase)) return "global::System.Guid";
        return reference?.TargetTypeId is not null && qualifiedTypeById.TryGetValue(reference.TargetTypeId, out var qualifiedName)
            ? qualifiedName + (reference.IsArray && !qualifiedName.EndsWith("[]", StringComparison.Ordinal) ? "[]" : "")
            : fallback ?? reference?.Name ?? "object";
    }
    private static IReadOnlyList<int> OverloadArities(DataMember member, bool isIndexer)
    {
        var count = member.Parameters.Count;
        if (count == 0) return new[] { 0 };
        var required = count;
        while (required > 0 && member.Parameters[required - 1].IsOptional) required--;
        if (isIndexer) required = Math.Max(1, required);
        return Enumerable.Range(required, count - required + 1).Reverse().ToArray();
    }

    private static IReadOnlyList<ProjectedParameter> Parameters(DataMember member, int arity, ProjectionNameRules rules, IReadOnlyDictionary<string, string> qualifiedTypeById, bool isIndexer, bool isEventSink)
    {
        if (member.Parameters.Count != 0)
            return member.Parameters.Take(arity).Select((parameter, index) => new ProjectedParameter
            {
                Name = parameter.Name,
                CSharpName = Sanitize(parameter.Name, rules),
                Type = QualifyBclType(QualifiedReferenceType(parameter.TypeReference, parameter.Type, qualifiedTypeById)),
                RefKind = isIndexer ? "value" : parameter.RefKind,
                IsOptional = parameter.IsOptional,
                HasDefaultValue = parameter.HasDefaultValue,
                DefaultValue = parameter.DefaultValue,
                EmitDefaultValue = false,
                Position = index,
                TypeReference = ProjectReference(parameter.TypeReference, qualifiedTypeById),
                SinkArgument = isEventSink ? EventSinkArgument(parameter.TypeReference, QualifyBclType(QualifiedReferenceType(parameter.TypeReference, parameter.Type, qualifiedTypeById)), Sanitize(parameter.Name, rules), parameter.RefKind) : null
            }).ToArray();
        return member.ParameterTypes.Take(arity).Select((type, index) => new ProjectedParameter { Name = "arg" + index.ToString(CultureInfo.InvariantCulture), CSharpName = "arg" + index.ToString(CultureInfo.InvariantCulture), Type = QualifyBclType(type), Position = index }).ToArray();
    }

    private static ProjectedParameter[] ApplyContractSinkArguments(ProjectedParameter[] parameters, IReadOnlyList<string>? attributes)
    {
        if (attributes is null || parameters.Length == 0) return parameters;
        var plans = new Dictionary<string, EventSinkArgumentPlan>(StringComparer.Ordinal);
        foreach (var attribute in attributes)
        {
            var match = SinkArgumentRegex.Match(attribute);
            if (!match.Success) continue;
            var name = match.Groups["name"].Value;
            var kind = match.Groups["kind"].Value;
            var explicitType = match.Groups["type"].Success ? match.Groups["type"].Value.Trim() : match.Groups["enum"].Value.Trim();
            ProjectionEventConversion conversion;
            string managedType;
            string? expression;
            if (kind.StartsWith("typeof(", StringComparison.Ordinal))
            {
                conversion = ProjectionEventConversion.KnownReference;
                managedType = explicitType;
                expression = null;
            }
            else
            {
                var scalarKind = kind["SinkArgumentType.".Length..];
                if (scalarKind.Equals("Enum", StringComparison.Ordinal))
                {
                    conversion = ProjectionEventConversion.Enum;
                    managedType = explicitType;
                    expression = "(" + managedType + ")global::System.Convert.ToInt32({0})";
                }
                else if (scalarKind.Equals("UnknownProxy", StringComparison.Ordinal))
                {
                    conversion = ProjectionEventConversion.EventReference;
                    managedType = "object";
                    expression = null;
                }
                else
                {
                    conversion = ProjectionEventConversion.Scalar;
                    managedType = scalarKind.Equals("String", StringComparison.Ordinal) ? "string"
                        : scalarKind.Equals("Bool", StringComparison.Ordinal) ? "bool"
                        : scalarKind;
                    var conversionName = scalarKind.Equals("Bool", StringComparison.Ordinal) ? "Boolean" : scalarKind;
                    expression = "global::System.Convert.To" + conversionName + "({0})";
                }
            }
            plans[name] = new EventSinkArgumentPlan
            {
                Conversion = conversion,
                ManagedType = managedType,
                WrapperTypeExpression = conversion == ProjectionEventConversion.KnownReference ? managedType + ".LateBindingApiWrapperType" : null,
                ConversionExpression = expression,
                WriteBackExpression = parameters.FirstOrDefault(parameter => parameter.CSharpName.Equals(name, StringComparison.Ordinal))?.RefKind is "ref" or "out" ? expression : null,
                IsContractDeclared = true
            };
        }
        return parameters.Select(parameter =>
        {
            if (!plans.TryGetValue(parameter.CSharpName, out var plan) && !plans.TryGetValue(parameter.Name, out plan))
                return parameter;
            if (parameter.SinkArgument is null || parameter.SinkArgument.Conversion == ProjectionEventConversion.Raw)
                return parameter with { SinkArgument = plan };
            return EquivalentSinkConversion(parameter.SinkArgument, plan)
                ? parameter with { SinkArgument = parameter.SinkArgument with { IsContractDeclared = true } }
                : parameter;
        }).ToArray();
    }

    private static ProjectedParameter[] ApplyContractSinkInvocations(ProjectedParameter[] parameters, IReadOnlyList<string>? invocationText)
    {
        if (invocationText is null || parameters.Length == 0) return parameters;
        var declarations = invocationText
            .Select(line => SinkLocalDeclarationRegex.Match(line))
            .Where(match => match.Success)
            .ToArray();
        return parameters.Select(parameter =>
        {
            var expectedLocalName = "new" + parameter.CSharpName;
            var declaration = declarations.FirstOrDefault(match =>
                match.Groups["local"].Value.Equals(expectedLocalName, StringComparison.OrdinalIgnoreCase));
            if (declaration is null)
            {
                declaration = declarations.FirstOrDefault(match =>
                    Regex.IsMatch(match.Groups["expression"].Value, @"(?<![A-Za-z0-9_])" + Regex.Escape(parameter.CSharpName) + @"(?![A-Za-z0-9_])", RegexOptions.CultureInvariant)
                    && !parameters.Any(output => match.Groups["local"].Value.Equals("new" + output.CSharpName, StringComparison.OrdinalIgnoreCase)));
            }
            if (declaration is null)
                return parameter with
                {
                    SinkArgument = new EventSinkArgumentPlan
                    {
                        Conversion = ProjectionEventConversion.Raw,
                        ManagedType = parameter.Type,
                        IsContractDeclared = true
                    }
                };

            var expression = declaration.Groups["expression"].Value;
            var managedType = declaration.Groups["type"].Value;
            var sourceArgument = parameters
                .Select(static input => input.CSharpName)
                .FirstOrDefault(input => Regex.IsMatch(expression, @"(?<![A-Za-z0-9_])" + Regex.Escape(input) + @"(?![A-Za-z0-9_])", RegexOptions.CultureInvariant));
            var localName = declaration.Groups["local"].Value;
            if (expression.Contains("CreateEventArgumentObjectFromComProxy", StringComparison.Ordinal))
                return parameter with
                {
                    SinkArgument = (parameter.SinkArgument ?? new EventSinkArgumentPlan()) with
                    {
                        Conversion = ProjectionEventConversion.EventReference,
                        ManagedType = managedType,
                        WrapperTypeExpression = null,
                        SourceArgument = sourceArgument,
                        ConversionExpression = null,
                        LocalName = localName,
                        IsContractDeclared = true
                    }
                };
            if (expression.Contains("CreateKnownObjectFromComProxy", StringComparison.Ordinal))
                return parameter with
                {
                    SinkArgument = (parameter.SinkArgument ?? new EventSinkArgumentPlan()) with
                    {
                        Conversion = ProjectionEventConversion.KnownReference,
                        ManagedType = managedType,
                        WrapperTypeExpression = managedType + ".LateBindingApiWrapperType",
                        SourceArgument = sourceArgument,
                        ConversionExpression = null,
                        LocalName = localName,
                        IsContractDeclared = true
                    }
                };
            if (!expression.Contains("Convert.To", StringComparison.Ordinal)) return parameter;
            var isEnum = expression.TrimStart().StartsWith("(", StringComparison.Ordinal);
            return parameter with
            {
                SinkArgument = (parameter.SinkArgument ?? new EventSinkArgumentPlan()) with
                {
                    Conversion = isEnum ? ProjectionEventConversion.Enum : ProjectionEventConversion.Scalar,
                    ManagedType = managedType,
                    WrapperTypeExpression = null,
                    ConversionExpression = sourceArgument is null
                        ? expression
                        : Regex.Replace(expression, @"(?<![A-Za-z0-9_])" + Regex.Escape(sourceArgument) + @"(?![A-Za-z0-9_])", "{0}", RegexOptions.CultureInvariant),
                    LocalName = localName,
                    SourceArgument = sourceArgument,
                    IsContractDeclared = true
                }
            };
        }).ToArray();
    }

    private static ContractMember OverlayEventSinkInvocation(ContractMember metadata, WrapperContract contract)
    {
        var parameterNames = ParseContractParameters(metadata.Parameters).Select(static parameter => parameter.CSharpName).ToArray();
        var sinkMember = contract.Types
            .Where(static type => type.Name.EndsWith("_SinkHelper", StringComparison.Ordinal))
            .SelectMany(static type => type.Members)
            .Where(candidate => candidate.Name.Equals(metadata.Name, StringComparison.Ordinal)
                                && candidate.Source.Equals(metadata.Source, StringComparison.Ordinal)
                                && ParseContractParameters(candidate.Parameters).Select(static parameter => parameter.CSharpName)
                                    .SequenceEqual(parameterNames, StringComparer.Ordinal))
            .OrderByDescending(static candidate => candidate.InvocationText.Count)
            .ThenBy(static candidate => candidate.Source, StringComparer.Ordinal)
            .ThenBy(static candidate => candidate.Line)
            .FirstOrDefault();
        return sinkMember is null
            ? metadata
            : metadata with
            {
                InvocationText = metadata.InvocationText.Concat(sinkMember.InvocationText)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(static line => line, StringComparer.Ordinal)
                    .ToArray()
            };
    }

    private static bool EquivalentSinkConversion(EventSinkArgumentPlan left, EventSinkArgumentPlan right)
        => left.Conversion == right.Conversion
           && TypeIdentity(left.ManagedType).Equals(TypeIdentity(right.ManagedType), StringComparison.OrdinalIgnoreCase);

    private static EventSinkArgumentPlan? OverlaySinkArgument(EventSinkArgumentPlan? sink, ProjectedParameter parameter)
    {
        if (sink is null) return null;
        var managedType = parameter.Type.Equals("object", StringComparison.Ordinal) ? sink.ManagedType : parameter.Type;
        return sink with
        {
            ManagedType = managedType,
            WrapperTypeExpression = sink.Conversion == ProjectionEventConversion.KnownReference ? managedType + ".LateBindingApiWrapperType" : null,
            WriteBackExpression = parameter.RefKind == "value" ? null : sink.ConversionExpression
        };
    }

    private static EventSinkArgumentPlan EventSinkArgument(DataTypeReference? reference, string managedType, string name, string refKind)
    {
        var identity = TypeIdentity(managedType);
        var conversion = reference?.IsEnum == true ? ProjectionEventConversion.Enum
            : reference?.IsComProxy == true && !string.IsNullOrWhiteSpace(reference.TypeKey) ? ProjectionEventConversion.KnownReference
            : reference?.IsComProxy == true ? ProjectionEventConversion.EventReference
            : ScalarTypes.Contains(identity) && ScalarEventExpression(identity) is not null ? ProjectionEventConversion.Scalar
            : ProjectionEventConversion.Raw;
        var expression = conversion == ProjectionEventConversion.Scalar ? ScalarEventExpression(identity) : null;
        return new EventSinkArgumentPlan
        {
            Conversion = conversion,
            ManagedType = managedType,
            WrapperTypeExpression = conversion == ProjectionEventConversion.KnownReference ? managedType + ".LateBindingApiWrapperType" : null,
            ConversionExpression = expression,
            WriteBackExpression = refKind == "value" ? null : expression
        };
    }

    private static string? ScalarEventExpression(string type)
    {
        var conversion = type.ToLowerInvariant() switch
        {
            "bool" or "boolean" => "Boolean",
            "byte" => "Byte",
            "sbyte" => "SByte",
            "short" or "int16" => "Int16",
            "ushort" or "uint16" => "UInt16",
            "int" or "int32" => "Int32",
            "uint" or "uint32" => "UInt32",
            "long" or "int64" => "Int64",
            "ulong" or "uint64" => "UInt64",
            "float" or "single" => "Single",
            "double" => "Double",
            "decimal" => "Decimal",
            "char" => "Char",
            "string" => "String",
            "datetime" => "DateTime",
            "intptr" => "IntPtr",
            "uintptr" => "UIntPtr",
            _ => type
        };
        return conversion is "IntPtr" or "UIntPtr" ? null : "global::System.Convert.To" + conversion + "({0})";
    }

    private static string FormatParameters(IReadOnlyList<ProjectedParameter> parameters)
        => "(" + string.Join(", ", parameters.Select(static parameter => (parameter.RefKind == "value" ? "" : parameter.RefKind + " ") + parameter.Type + " " + parameter.CSharpName)) + ")";

    private static string NormalizeMemberKind(string kind, string? accessorKind, int parameterCount, bool isIndexer)
    {
        if (kind.Contains("constructor", StringComparison.OrdinalIgnoreCase) || kind.Equals("ctor", StringComparison.OrdinalIgnoreCase)) return "constructor";
        if (kind.Contains("event", StringComparison.OrdinalIgnoreCase)) return "event";
        if (kind.Contains("field", StringComparison.OrdinalIgnoreCase) || kind.Contains("constant", StringComparison.OrdinalIgnoreCase)) return "field";
        if (isIndexer) return "indexer";
        if (accessorKind is not null || kind.Contains("property", StringComparison.OrdinalIgnoreCase))
            return parameterCount == 0 ? "property" : "method";
        return "method";
    }

    private static ContractMemberMatch ChooseContractMember(ContractType? primary, IReadOnlyList<ContractType> auxiliaryTypes, IReadOnlyList<ContractType> inheritedTypes, DataMember data, int arity, bool isIndexer)
    {
        var ownerLogicalId = primary?.LogicalId ?? data.TypeId;
        var member = SelectContractMember(primary?.Members ?? Array.Empty<ContractMember>(), data, arity, isIndexer);
        if (member is not null) return new ContractMemberMatch(member, ownerLogicalId, ProjectionMemberEmissionDisposition.Emit);
        foreach (var auxiliary in auxiliaryTypes)
        {
            member = SelectContractMember(auxiliary.Members, data, arity, isIndexer);
            if (member is not null) return new ContractMemberMatch(member, auxiliary.LogicalId, ProjectionMemberEmissionDisposition.Emit);
        }
        foreach (var inherited in inheritedTypes)
        {
            member = SelectContractMember(inherited.Members, data, arity, isIndexer);
            if (member is not null) return new ContractMemberMatch(member, inherited.LogicalId, ProjectionMemberEmissionDisposition.Inherited);
        }
        return new ContractMemberMatch(null, ownerLogicalId, ProjectionMemberEmissionDisposition.Emit);
    }

    private static ContractMember? SelectContractMember(IReadOnlyList<ContractMember> members, DataMember data, int arity, bool isIndexer)
    {
        var matches = members.Where(candidate => (string.Equals(candidate.Name, data.Name, StringComparison.Ordinal) || isIndexer && candidate.Kind.Equals("indexer", StringComparison.OrdinalIgnoreCase)) && ParameterCount(candidate.Parameters) == arity).ToArray();
        if (matches.Length == 0) return null;
        return matches
            .OrderByDescending(candidate => ContractMemberScore(candidate, data, arity))
            .ThenBy(static candidate => candidate.Source, StringComparer.Ordinal)
            .ThenBy(static candidate => candidate.Line)
            .ThenBy(static candidate => candidate.Signature, StringComparer.Ordinal)
            .First();
    }

    private static int ContractMemberScore(ContractMember candidate, DataMember data, int arity)
    {
        var score = 0;
        var candidateProperty = candidate.Kind.Contains("property", StringComparison.OrdinalIgnoreCase);
        if (data.AccessorKind is null ? !candidateProperty : candidateProperty) score += 100;
        if (candidate.ReturnType is not null && data.ReturnType is not null
            && TypeIdentity(candidate.ReturnType).Equals(TypeIdentity(data.ReturnType), StringComparison.OrdinalIgnoreCase)) score += 20;
        var contractParameters = ParseContractParameters(candidate.Parameters);
        var dataParameters = data.Parameters.Take(arity).ToArray();
        for (var index = 0; index < Math.Min(contractParameters.Count, dataParameters.Length); index++)
            if (TypeIdentity(contractParameters[index].Type).Equals(TypeIdentity(dataParameters[index].Type), StringComparison.OrdinalIgnoreCase)) score += 10;
        return score;
    }

    private static string TypeIdentity(string value)
    {
        var normalized = value.Replace("global::", "", StringComparison.Ordinal).Trim().TrimEnd('?');
        var generic = normalized.IndexOf('<');
        var prefix = generic < 0 ? normalized : normalized[..generic];
        return prefix.Split('.').Last() + (generic < 0 ? "" : normalized[generic..]);
    }

    private static IReadOnlyList<ProjectedEventBinding> ProjectEventBindings(DataType dataType, ContractType? contract, WrapperContract wrapperContract, IReadOnlyDictionary<string, DataType> typeById, IReadOnlyDictionary<string, string> qualifiedTypeById)
    {
        var contractSinkTypes = (contract?.Attributes ?? Array.Empty<string>())
            .Where(static attribute => attribute.StartsWith("EventSink(", StringComparison.Ordinal))
            .SelectMany(attribute => Regex.Matches(attribute, @"typeof\((?<type>[^)]+)\)", RegexOptions.CultureInvariant).Select(static match => match.Groups["type"].Value.Trim()))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (contractSinkTypes.Length != 0)
            return contractSinkTypes.Select((sinkType, ordinal) => new ProjectedEventBinding
            {
                LogicalId = (contract?.LogicalId ?? dataType.LogicalId) + "/event-binding/" + ordinal.ToString(CultureInfo.InvariantCulture),
                SinkHelperType = sinkType,
                FieldName = EventSinkFieldName(sinkType)
            }).ToArray();
        return dataType.EventInterfaceIds.Where(typeById.ContainsKey)
            .Where(id => wrapperContract.Types.Any(contractType => contractType.Name.Equals(typeById[id].Name, StringComparison.Ordinal)))
            .Select((id, ordinal) =>
        {
            var eventType = typeById[id];
            var sinkType = qualifiedTypeById.GetValueOrDefault(id, eventType.Name) + "_SinkHelper";
            return new ProjectedEventBinding
            {
                LogicalId = dataType.LogicalId + "/event-binding/" + ordinal.ToString(CultureInfo.InvariantCulture),
                SinkHelperType = sinkType,
                FieldName = EventSinkFieldName(sinkType),
                EventInterfaceLogicalId = id,
                InterfaceId = eventType.DeclaredGuid
            };
        }).ToArray();
    }

    private static string EventSinkFieldName(string sinkType)
    {
        var simple = sinkType.Split('.').Last();
        var stem = simple.EndsWith("_SinkHelper", StringComparison.Ordinal) ? simple[..^"_SinkHelper".Length] : simple;
        stem = stem.TrimStart('_');
        return "_" + (stem.Length == 0 ? "event" : char.ToLowerInvariant(stem[0]) + stem[1..]) + "_SinkHelper";
    }

    private static IReadOnlyList<VersionSupport> SupportFor(string targetId, IReadOnlyDictionary<string, SupportObservation[]> observations)
    {
        if (!observations.TryGetValue(targetId, out var records)) return Array.Empty<VersionSupport>();
        return records.GroupBy(static item => item.Product, StringComparer.Ordinal).OrderBy(static group => group.Key, StringComparer.Ordinal).Select(static group => new VersionSupport
        {
            Product = group.Key,
            Versions = group.SelectMany(static item => item.Versions).Distinct(StringComparer.Ordinal).OrderBy(static item => item, VersionComparer.Instance).ToArray()
        }).ToArray();
    }

    private static IReadOnlyList<VersionSupport> ParseSupport(IEnumerable<string> attributes)
    {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var attribute in attributes)
            foreach (Match match in SupportRegex.Matches(attribute))
            {
                var product = match.Groups["product"].Value;
                if (!result.TryGetValue(product, out var versions)) result[product] = versions = new HashSet<string>(StringComparer.Ordinal);
                foreach (var version in match.Groups["versions"].Value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) versions.Add(version.Trim('"'));
            }
        return result.OrderBy(static pair => pair.Key, StringComparer.Ordinal).Select(static pair => new VersionSupport { Product = pair.Key, Versions = pair.Value.OrderBy(static item => item, VersionComparer.Instance).ToArray() }).ToArray();
    }

    private static IReadOnlyList<VersionSupport> MergeSupport(params IReadOnlyList<VersionSupport>[] sources)
        => sources.SelectMany(static source => source).GroupBy(static item => item.Product, StringComparer.Ordinal).OrderBy(static group => group.Key, StringComparer.Ordinal).Select(static group => new VersionSupport { Product = group.Key, Versions = group.SelectMany(static item => item.Versions).Distinct(StringComparer.Ordinal).OrderBy(static item => item, VersionComparer.Instance).ToArray() }).ToArray();

    private static IReadOnlyList<string> ContractOrDataAttributes(IReadOnlyList<string>? contractAttributes, IReadOnlyList<VersionSupport> support, ProjectionEntityKind kind)
    {
        var result = new HashSet<string>((contractAttributes ?? Array.Empty<string>()).Select(NormalizeAttribute), StringComparer.Ordinal);
        AddSupportAttributes(result, support);
        if (kind == ProjectionEntityKind.EventInterface)
            result.Add("InternalEntity(InternalEntityKind.ComEventInterface)");
        else if (!result.Any(static attribute => attribute.Contains("EntityType", StringComparison.Ordinal)))
            result.Add("EntityType(EntityType." + EntityTypeAttributeValue(kind) + ")");
        return result.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static string EntityTypeAttributeValue(ProjectionEntityKind kind)
        => kind switch
        {
            ProjectionEntityKind.DispatchInterface => "IsDispatchInterface",
            ProjectionEntityKind.Interface => "IsInterface",
            ProjectionEntityKind.CoClass => "IsCoClass",
            ProjectionEntityKind.Enum => "IsEnum",
            ProjectionEntityKind.Module => "IsModule",
            ProjectionEntityKind.Constants => "IsConstants",
            ProjectionEntityKind.Record or ProjectionEntityKind.TypeDef => "IsStruct",
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Event interfaces use InternalEntity instead of EntityType.")
        };

    private static IReadOnlyList<string> ContractOrSupportAttributes(IReadOnlyList<string>? contractAttributes, IReadOnlyList<VersionSupport> support)
    {
        var result = new HashSet<string>((contractAttributes ?? Array.Empty<string>()).Select(NormalizeAttribute), StringComparer.Ordinal);
        AddSupportAttributes(result, support);
        return result.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<string> DataMemberAttributes(IReadOnlyList<string>? contractAttributes, IReadOnlyList<VersionSupport> support, DataMember member)
    {
        var result = new HashSet<string>(ContractOrSupportAttributes(contractAttributes, support), StringComparer.Ordinal);
        if (member.AccessorKind is not ("put" or "putref") && (member.IsComProxy || member.ReturnTypeReference?.IsComProxy == true)) result.Add("ProxyResult");
        if (member.IsHidden)
        {
            result.Add("EditorBrowsable(EditorBrowsableState.Never)");
            result.Add("Browsable(false)");
        }
        return result.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static string NormalizeAttribute(string attribute)
    {
        if (string.IsNullOrWhiteSpace(attribute)) return attribute;
        return attribute
            .Replace("ComInterfaceType.", "System.Runtime.InteropServices.ComInterfaceType.", StringComparison.Ordinal)
            .Replace("TypeLibTypeFlags.", "System.Runtime.InteropServices.TypeLibTypeFlags.", StringComparison.Ordinal)
            .Replace("ClassInterfaceType.", "System.Runtime.InteropServices.ClassInterfaceType.", StringComparison.Ordinal)
            .Replace("UnmanagedType.", "System.Runtime.InteropServices.UnmanagedType.", StringComparison.Ordinal)
            .Replace("MethodImplOptions.", "System.Runtime.CompilerServices.MethodImplOptions.", StringComparison.Ordinal)
            .Replace("MethodCodeType.", "System.Runtime.CompilerServices.MethodCodeType.", StringComparison.Ordinal)
            .Replace("ComImport", "System.Runtime.InteropServices.ComImport", StringComparison.Ordinal)
            .Replace("ComVisible(", "System.Runtime.InteropServices.ComVisible(", StringComparison.Ordinal)
            .Replace("Guid(", "System.Runtime.InteropServices.Guid(", StringComparison.Ordinal)
            .Replace("InterfaceType(", "System.Runtime.InteropServices.InterfaceType(", StringComparison.Ordinal)
            .Replace("TypeLibType(", "System.Runtime.InteropServices.TypeLibType(", StringComparison.Ordinal)
            .Replace("ClassInterface(", "System.Runtime.InteropServices.ClassInterface(", StringComparison.Ordinal)
            .Replace("MarshalAs(", "System.Runtime.InteropServices.MarshalAs(", StringComparison.Ordinal)
            .Replace("DispId(", "System.Runtime.InteropServices.DispId(", StringComparison.Ordinal)
            .Replace("MethodImpl(", "System.Runtime.CompilerServices.MethodImpl(", StringComparison.Ordinal);
    }

    private static void AddSupportAttributes(ISet<string> result, IReadOnlyList<VersionSupport> support)
    {
        foreach (var item in support)
        {
            var prefix = "SupportByVersion(\"" + item.Product + "\"";
            if (!result.Any(attribute => attribute.StartsWith(prefix, StringComparison.Ordinal))) result.Add(prefix + (item.Versions.Count == 0 ? "" : ", " + string.Join(",", item.Versions)) + ")");
        }
    }

    private static IReadOnlyList<string> TypeCapabilities(DataType type, IReadOnlyList<DataMember> members, TypeDescriptor descriptor, ContractType? contract, ProjectionPolicy policy)
    {
        var capabilities = new HashSet<string>(StringComparer.Ordinal);
        if (type.IsEventInterface == true || type.EventInterfaceIds.Count != 0 || members.Any(static member => member.Kind.Contains("event", StringComparison.OrdinalIgnoreCase))) capabilities.Add("event");
        if (members.Any(static member => member.Name.Equals("Item", StringComparison.Ordinal) && member.Parameters.Count != 0)) capabilities.Add("indexer");
        if (members.Any(static member => member.Name.Contains("GetEnumerator", StringComparison.OrdinalIgnoreCase) || member.DispId == -4)) capabilities.Add("enumerator");
        if (capabilities.Contains("enumerator") || (members.Any(static member => member.Name.Equals("Item", StringComparison.Ordinal)) && members.Any(static member => member.Name.Equals("Count", StringComparison.Ordinal)))) capabilities.Add("collection");
        if (contract is not null)
        {
            if (contract.Attributes.Any(static attribute => attribute.Contains("HasIndexProperty", StringComparison.OrdinalIgnoreCase))) capabilities.Add("indexer");
            if (contract.Interfaces.Any(static item => item.Contains("IEnumerable", StringComparison.OrdinalIgnoreCase))
                || contract.Attributes.Any(static attribute => attribute.Contains("Enumerator(", StringComparison.OrdinalIgnoreCase)))
            {
                capabilities.Add("enumerator");
                capabilities.Add("collection");
            }
            if (contract.BaseType?.EndsWith("Collection", StringComparison.Ordinal) == true) capabilities.Add("collection");
            if (contract.Members.Any(static member => member.Kind.Contains("event", StringComparison.OrdinalIgnoreCase))) capabilities.Add("event");
        }
        if (descriptor.EntityKind == ProjectionEntityKind.CoClass) capabilities.Add("activation");
        return AddPolicyCapabilities(capabilities, "type", policy);
    }

    private static IReadOnlyList<string> MemberCapabilities(DataMember data, ContractMember? member, string kind, InvocationPlan invocation, ProjectionPolicy policy)
    {
        var capabilities = new HashSet<string>(StringComparer.Ordinal);
        if (kind == "event") capabilities.Add("event");
        if (kind == "indexer") capabilities.Add("indexer");
        if (data.Name.Contains("GetEnumerator", StringComparison.OrdinalIgnoreCase) || data.DispId == -4 || (data.ReturnType?.Contains("IEnumerator", StringComparison.OrdinalIgnoreCase) ?? false)) capabilities.Add("enumerator");
        if (invocation.ReturnConversion is ProjectionReturnConversion.KnownReference or ProjectionReturnConversion.Reference) capabilities.Add("proxy-result");
        if ((member?.Attributes ?? Array.Empty<string>()).Any(static attribute => attribute.Contains("IndexProperty", StringComparison.OrdinalIgnoreCase))) capabilities.Add("indexer");
        return AddPolicyCapabilities(capabilities, kind, policy);
    }

    private static IReadOnlyList<string> AddPolicyCapabilities(HashSet<string> capabilities, string key, ProjectionPolicy policy)
    {
        if (policy.RuntimeCapabilities.TryGetValue(key, out var configured))
            foreach (var item in configured.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) capabilities.Add(item);
        return capabilities.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<string> TypeRuntimeRequirements(TypeDescriptor descriptor, IReadOnlyList<string> capabilities, IReadOnlyList<string> attributes, ProjectionPolicy policy)
    {
        var result = new HashSet<string>(RuntimeRequirements("type", descriptor.EntityKind.ToString(), attributes, policy), StringComparer.Ordinal);
        if (descriptor.CSharpKind == "class" && descriptor.EntityKind is not ProjectionEntityKind.Module and not ProjectionEntityKind.Constants)
        {
            result.Add("netoffice-core");
            result.Add("com-proxy");
            result.Add("wrapper-type-cache");
        }
        if (capabilities.Contains("collection", StringComparer.Ordinal)) result.Add("collection-runtime");
        if (capabilities.Contains("enumerator", StringComparer.Ordinal)) result.Add("enumerator-runtime");
        if (capabilities.Contains("event", StringComparer.Ordinal)) result.Add("event-sink-runtime");
        if (descriptor.EntityKind == ProjectionEntityKind.CoClass) result.Add("com-activation");
        return result.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<string> RuntimeRequirements(string scope, string kind, IReadOnlyList<string> attributes, ProjectionPolicy policy)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (policy.RuntimeCapabilities.TryGetValue(scope + ":" + kind, out var configured))
            foreach (var item in configured.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) result.Add(item);
        foreach (var attribute in attributes.SelectMany(static item => AttributeRegex.Matches(item).Select(static match => match.Value)).OrderBy(static item => item, StringComparer.Ordinal))
            if (policy.RuntimeCapabilities.TryGetValue("attribute:" + attribute, out var requirement)) result.Add(requirement);
        return result.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<ProjectedContractMember> ProjectExactContractMembers(
        ContractType? contract,
        ProjectionPolicy policy,
        IReadOnlyList<ProjectedMember>? projectedMembers = null)
    {
        if (contract is null) return Array.Empty<ProjectedContractMember>();
        var localForwardNames = projectedMembers?
            .Where(static member => member.Invocation.OperationKind == ProjectionInvocationOperation.LocalForward)
            .Select(static member => member.CSharpName)
            .ToHashSet(StringComparer.Ordinal) ?? EmptyStringSet;
        return contract.Members
            .Where(member => RetainedExactContractMemberNames.Contains(member.Name)
                             || member.Name.EndsWith(".GetEnumerator", StringComparison.Ordinal)
                             || localForwardNames.Contains(member.Name))
            .OrderBy(static member => member.Name, StringComparer.Ordinal)
            .ThenBy(static member => member.Signature, StringComparer.Ordinal)
            .ThenBy(static member => member.Source, StringComparer.Ordinal)
            .ThenBy(static member => member.Line)
            .Select((member, ordinal) =>
            {
                var logicalId = contract.LogicalId + "/contract-facet/" + ordinal.ToString(CultureInfo.InvariantCulture);
                return new ProjectedContractMember
                {
                    LogicalId = logicalId,
                    Name = member.Name,
                    Kind = member.Kind,
                    Accessibility = member.Accessibility,
                    Modifiers = member.Modifiers,
                    ReturnType = member.ReturnType,
                    Parameters = member.Parameters,
                    ParameterList = ParseContractParameters(member.Parameters),
                    Signature = member.Signature,
                    Attributes = member.Attributes.Select(NormalizeAttribute).ToArray(),
                    SupportVersions = ParseSupport(member.Attributes),
                    DocsBindingKey = DocsKey(policy, logicalId),
                    Source = member.Source,
                    Line = member.Line
                };
            }).ToArray();
    }

    private static IReadOnlyList<ConstructorPlan> ConstructorPlans(TypeDescriptor descriptor, string product, string typeName, ContractType? contract = null, ProjectionPolicy? policy = null)
    {
        if (descriptor.CSharpKind != "class" || descriptor.EntityKind is ProjectionEntityKind.Module or ProjectionEntityKind.Constants) return Array.Empty<ConstructorPlan>();
        var contractConstructors = contract?.Members.Where(static member => member.Kind.Equals("constructor", StringComparison.OrdinalIgnoreCase))
            .OrderBy(static member => member.Signature, StringComparer.Ordinal).ThenBy(static member => member.Source, StringComparer.Ordinal).ThenBy(static member => member.Line).ToArray()
            ?? Array.Empty<ContractMember>();
        if (contractConstructors.Length != 0 && policy is not null)
            return contractConstructors.Select((member, ordinal) =>
            {
                var logicalId = contract!.LogicalId + "/constructor/" + ordinal.ToString(CultureInfo.InvariantCulture);
                return new ConstructorPlan
                {
                    Kind = ConstructorKind(member),
                    LogicalId = logicalId,
                    IsContractOverlay = true,
                    Signature = member.Signature,
                    Accessibility = member.Accessibility,
                    Attributes = member.Attributes.Select(NormalizeAttribute).ToArray(),
                    Parameters = ParseContractParameters(member.Parameters),
                    BaseCall = ConstructorBaseCall(member.Signature),
                    RuntimeRequirements = new[] { "com-proxy" },
                    DocsBindingKey = DocsKey(policy, logicalId),
                    Source = member.Source,
                    Line = member.Line
                };
            }).ToArray();
        var defaultBaseCall = descriptor.EntityKind == ProjectionEntityKind.CoClass
            ? "base(" + System.Text.Json.JsonSerializer.Serialize(product + "." + typeName) + ")"
            : "base()";
        return new[]
        {
            Constructor("proxy-share", "base(factory, parentObject, proxyShare)", Param("factory", "Core", 0), Param("parentObject", "ICOMObject", 1), Param("proxyShare", "COMProxyShare", 2)),
            Constructor("proxy", "base(factory, parentObject, comProxy)", Param("factory", "Core", 0), Param("parentObject", "ICOMObject", 1), Param("comProxy", "object", 2)),
            Constructor("proxy", "base(parentObject, comProxy)", Param("parentObject", "ICOMObject", 0), Param("comProxy", "object", 1)),
            Constructor("typed-proxy", "base(factory, parentObject, comProxy, comProxyType)", Param("factory", "Core", 0), Param("parentObject", "ICOMObject", 1), Param("comProxy", "object", 2), Param("comProxyType", "global::System.Type", 3)),
            Constructor("typed-proxy", "base(parentObject, comProxy, comProxyType)", Param("parentObject", "ICOMObject", 0), Param("comProxy", "object", 1), Param("comProxyType", "global::System.Type", 2)),
            Constructor("replacement", "base(replacedObject)", Param("replacedObject", "ICOMObject", 0)),
            Constructor(descriptor.EntityKind == ProjectionEntityKind.CoClass ? "activation" : "default", defaultBaseCall),
            Constructor("progid", "base(progId)", Param("progId", "string", 0))
        };
    }

    private static string ActivationProgId(IReadOnlyList<ConstructorPlan> constructors, string fallback)
    {
        foreach (var constructor in constructors)
        {
            var match = Regex.Match(constructor.BaseCall, "^base\\((?<literal>\\\"(?:\\\\.|[^\\\"])*\\\")\\)$", RegexOptions.CultureInvariant);
            if (match.Success)
                return System.Text.Json.JsonSerializer.Deserialize<string>(match.Groups["literal"].Value) ?? fallback;
        }
        return fallback;
    }

    private static string ConstructorKind(ContractMember member)
    {
        var parameters = ParseContractParameters(member.Parameters);
        if (parameters.Count == 0) return "default";
        if (parameters.Count == 1 && parameters[0].Type.Equals("string", StringComparison.Ordinal)) return "progid";
        if (parameters.Count == 1 && parameters[0].Type.EndsWith("ICOMObject", StringComparison.Ordinal)) return "replacement";
        return parameters.Any(static parameter => parameter.Type.EndsWith("Type", StringComparison.Ordinal)) ? "typed-proxy" : "proxy";
    }

    private static ProjectionRuntimeMemberKind ContractRuntimeMemberKind(ContractMember member)
    {
        if (member.InvocationText.Count != 0) return ProjectionRuntimeMemberKind.None;
        if (member.Name.Equals("Clone", StringComparison.Ordinal)
            && member.Modifiers.Contains("new", StringComparer.Ordinal)
            && member.Modifiers.Contains("virtual", StringComparer.Ordinal))
            return ProjectionRuntimeMemberKind.Clone;
        if (member.Name.Equals("FromProxyService", StringComparison.Ordinal)
            && member.Kind.Equals("property", StringComparison.OrdinalIgnoreCase)
            && member.Attributes.Any(static attribute => attribute.Contains("EditorBrowsable(EditorBrowsableState.Advanced)", StringComparison.Ordinal)))
            return ProjectionRuntimeMemberKind.FromProxyService;
        if (member.Name.Equals("Dispose", StringComparison.Ordinal)
            && member.Modifiers.Contains("override", StringComparer.Ordinal))
            return ParameterCount(member.Parameters) == 0
                ? ProjectionRuntimeMemberKind.Dispose
                : ProjectionRuntimeMemberKind.DisposeWithEventBinding;
        return ProjectionRuntimeMemberKind.None;
    }

    private static string ConstructorBaseCall(string signature)
    {
        var match = Regex.Match(signature, @":\s*(?<kind>base|this)\s*\(", RegexOptions.CultureInvariant);
        if (!match.Success) return "base()";
        var constructorKind = match.Groups["kind"].Value;
        var start = signature.IndexOf(constructorKind + "(", match.Index, StringComparison.Ordinal);
        var depth = 0;
        for (var index = start; index < signature.Length; index++)
        {
            if (signature[index] == '(') depth++;
            else if (signature[index] == ')' && --depth == 0) return signature[start..(index + 1)];
        }
        return "base()";
    }

    private static ConstructorPlan Constructor(string kind, string baseCall, params ProjectedParameter[] parameters)
        => new() { Kind = kind, BaseCall = baseCall, Parameters = parameters, RuntimeRequirements = new[] { kind is "progid" or "activation" ? "com-activation" : "com-proxy" } };
    private static ProjectedParameter Param(string name, string type, int position) => new() { Name = name, CSharpName = name, Type = type, Position = position };

    private static IReadOnlyList<string> ResolveBaseNames(DataType type, IReadOnlyDictionary<string, string> qualifiedTypeById, ContractType? contract)
    {
        if (contract is not null) return new[] { contract.BaseType }.Where(static item => !string.IsNullOrWhiteSpace(item)).Concat(contract.Interfaces).Cast<string>().Select(QualifyBclType).Distinct(StringComparer.Ordinal).ToArray();
        return type.BaseTypeIds.Concat(type.DefaultInterfaceIds).Select(id => qualifiedTypeById.GetValueOrDefault(id, id)).Distinct(StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<string> InterfaceNames(DataType type, IReadOnlyDictionary<string, string> qualifiedTypeById, string? baseType)
        => type.BaseTypeIds.Concat(type.DefaultInterfaceIds)
            .Concat(Describe(type).EntityKind == ProjectionEntityKind.CoClass ? Array.Empty<string>() : type.EventInterfaceIds)
            .Select(id => qualifiedTypeById.GetValueOrDefault(id, id))
            .Where(name => !string.Equals(name, baseType, StringComparison.Ordinal))
            .Distinct(StringComparer.Ordinal)
            .ToArray();

    private static IReadOnlyList<DataType> EffectiveBaseTypes(DataType type, IReadOnlyDictionary<string, DataType> types)
    {
        var result = new List<DataType>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        void Visit(DataType current)
        {
            foreach (var id in current.BaseTypeIds.Concat(current.DefaultInterfaceIds))
                if (seen.Add(id) && types.TryGetValue(id, out var baseType)) { result.Add(baseType); Visit(baseType); }
        }
        Visit(type);
        return result;
    }

    private static string QualifyTypeName(DataType type, IReadOnlyDictionary<string, DataLibrary> libraries, IReadOnlyDictionary<string, DataProject> projects)
    {
        projects.TryGetValue(type.ProjectId, out var project);
        if (!libraries.TryGetValue(type.LibraryId, out var library) && project is null) return type.Name;
        if (!string.IsNullOrWhiteSpace(type.Namespace) && !type.Namespace.StartsWith("NetOffice.", StringComparison.Ordinal))
            return type.Namespace + "." + type.Name;
        var descriptor = Describe(type);
        var root = project?.Namespace;
        if (string.IsNullOrWhiteSpace(root)) root = "NetOffice." + (project?.Name ?? library!.Name) + "Api";
        return NamespaceFromRoot(root, descriptor.Category) + "." + type.Name;
    }

    private static string? DefaultBaseType(TypeDescriptor descriptor, IReadOnlyList<string> capabilities)
    {
        if (descriptor.CSharpKind != "class" || descriptor.EntityKind is ProjectionEntityKind.Module or ProjectionEntityKind.Constants) return null;
        return "NetOffice.COMObject";
    }

    private static string QualifyBclType(string type)
        => type.StartsWith("System.", StringComparison.Ordinal) ? "global::" + type : type;

    private static string NormalizeTopLevelAccessibility(string? accessibility)
        => accessibility is "private" or "protected" or "internal" or "public" ? accessibility : "public";

    private static string Namespace(string product, ProjectionFileCategory category)
        => NamespaceFromRoot("NetOffice." + product + "Api", category);

    private static string NamespaceFromRoot(string root, ProjectionFileCategory category)
        => category switch
        {
            ProjectionFileCategory.Enums => root + ".Enums",
            ProjectionFileCategory.Constants => root + ".Constants",
            ProjectionFileCategory.Records => root + ".Records",
            ProjectionFileCategory.TypeDefs => root + ".TypeDefs",
            ProjectionFileCategory.Events => root + ".Events",
            _ => root
        };

    private static ProjectionFileCategory SourceCategory(string value, ProjectionFileCategory fallback)
        => value.Trim().ToLowerInvariant() switch
        {
            "dispatchinterfaces" => ProjectionFileCategory.DispatchInterfaces,
            "interfaces" => ProjectionFileCategory.Interfaces,
            "events" => ProjectionFileCategory.Events,
            "coclasses" or "classes" => ProjectionFileCategory.Classes,
            "enums" => ProjectionFileCategory.Enums,
            "modules" => ProjectionFileCategory.Modules,
            "constants" => ProjectionFileCategory.Constants,
            "records" => ProjectionFileCategory.Records,
            "typedefs" => ProjectionFileCategory.TypeDefs,
            _ => fallback
        };

    private static IReadOnlyList<string> TypeModifiers(TypeDescriptor descriptor, ContractType? contract)
    {
        var result = new HashSet<string>(contract?.Modifiers ?? Array.Empty<string>(), StringComparer.Ordinal);
        if (descriptor.EntityKind is ProjectionEntityKind.Module or ProjectionEntityKind.Constants) result.Add("static");
        return result.OrderBy(static item => item, StringComparer.Ordinal).ToArray();
    }

    private static string TypeSignature(TypeDescriptor descriptor, string name, string? baseType, IReadOnlyList<string> interfaces)
    {
        var modifier = descriptor.EntityKind is ProjectionEntityKind.Module or ProjectionEntityKind.Constants ? "static " : "";
        var bases = new[] { baseType }.Where(static item => !string.IsNullOrWhiteSpace(item)).Concat(interfaces).Cast<string>().ToArray();
        return "public " + modifier + descriptor.CSharpKind + " " + name + (bases.Length == 0 ? "" : " : " + string.Join(", ", bases));
    }

    private static string DataSignature(string kind, string? accessorKind, string? returnType, string name, string parameters)
    {
        var access = accessorKind == "get" ? "get;" : "set;";
        if (kind == "property") return $"public {returnType ?? "object"} {name} {{ {access} }}";
        if (kind == "indexer") return $"public {returnType ?? "object"} this[{parameters.Trim('(', ')')}] {{ {access} }}";
        return $"public {returnType ?? "void"} {name}{parameters}";
    }

    private static IReadOnlyList<WrapperFile> PartitionContractParts(ProjectedType type, WrapperContract contract, string primarySource, ProjectionPolicy policy)
    {
        static bool IsSourceFile(string? value) => !string.IsNullOrWhiteSpace(value) && value.EndsWith(".cs", StringComparison.OrdinalIgnoreCase);
        var sources = new[] { primarySource }
            .Concat(type.Members.Where(static member => member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit).Select(static member => member.Source))
            .Concat(type.Constructors.Select(static constructor => constructor.Source))
            .Concat(type.ExactContractMembers.Select(static member => member.Source))
            .Where(IsSourceFile)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(source => source.Equals(primarySource, StringComparison.Ordinal) ? 0 : 1)
            .ThenBy(static source => source, StringComparer.Ordinal)
            .ToArray();
        if (sources.Length <= 1)
        {
            var path = PartitionPath(type, primarySource, policy);
            return new[] { new WrapperFile { Path = path, Namespace = type.Namespace, Category = type.FileCategory, Types = new[] { type } } };
        }

        var result = new List<WrapperFile>(sources.Length);
        var knownSources = sources.ToHashSet(StringComparer.Ordinal);
        foreach (var source in sources)
        {
            var isPrimary = source.Equals(primarySource, StringComparison.Ordinal);
            var contractPart = contract.Types
                .Where(candidate => candidate.Name.Equals(type.Name, StringComparison.Ordinal)
                                    && (candidate.Source.Equals(source, StringComparison.Ordinal)
                                        || candidate.Members.Any(member => member.Source.Equals(source, StringComparison.Ordinal))))
                .OrderByDescending(candidate => candidate.Source.Equals(source, StringComparison.Ordinal))
                .ThenBy(static candidate => candidate.LogicalId, StringComparer.Ordinal)
                .FirstOrDefault();
            var partAccessibility = contractPart is null ? type.Accessibility : NormalizeTopLevelAccessibility(contractPart.Accessibility);
            if (isPrimary
                && contractPart is not null
                && !contractPart.Source.Equals(source, StringComparison.Ordinal)
                && Path.GetFileName(source).Equals(type.Name + ".cs", StringComparison.OrdinalIgnoreCase)
                && partAccessibility.Equals("private", StringComparison.Ordinal))
                partAccessibility = "public";
            var partId = isPrimary ? type.LogicalId
                : contractPart is not null && !contractPart.LogicalId.Equals(type.LogicalId, StringComparison.Ordinal) ? contractPart.LogicalId
                : type.CanonicalLogicalId + "/part/" + ShortHash(source);
            bool InPart(string? memberSource) => string.Equals(memberSource, source, StringComparison.Ordinal)
                                                  || isPrimary && (!IsSourceFile(memberSource) || !knownSources.Contains(memberSource!));
            var partMembers = type.Members.Where(member => InPart(member.Source)).Select(member =>
                member.EmissionOwnerLogicalId.Equals(type.LogicalId, StringComparison.Ordinal)
                    ? member with { EmissionOwnerLogicalId = partId }
                    : member).ToArray();
            var part = type with
            {
                LogicalId = partId,
                CanonicalLogicalId = type.CanonicalLogicalId,
                IsPrimaryPart = isPrimary,
                ContractPartLogicalId = contractPart?.LogicalId ?? type.ContractPartLogicalId,
                ContractPartSource = source,
                Namespace = contractPart?.Namespace ?? type.Namespace,
                Accessibility = partAccessibility,
                Modifiers = contractPart?.Modifiers ?? type.Modifiers,
                Attributes = contractPart is null ? type.Attributes : contractPart.Attributes.Select(NormalizeAttribute).ToArray(),
                BaseType = contractPart?.BaseType ?? type.BaseType,
                Interfaces = contractPart?.Interfaces ?? type.Interfaces,
                Signature = contractPart?.Signature ?? type.Signature,
                SupportVersions = contractPart is null ? type.SupportVersions : ParseSupport(contractPart.Attributes),
                EventSink = isPrimary ? type.EventSink : null,
                AuxiliaryTypes = type.AuxiliaryTypes.Where(auxiliary => InPart(auxiliary.Source)).ToArray(),
                EventBindings = isPrimary ? type.EventBindings : Array.Empty<ProjectedEventBinding>(),
                Constructors = type.Constructors.Where(constructor => InPart(constructor.Source)).ToArray(),
                ExactContractMembers = type.ExactContractMembers.Where(member => InPart(member.Source)).ToArray(),
                DocsBindingKey = isPrimary ? type.DocsBindingKey : DocsKey(policy, partId),
                Source = source,
                Line = contractPart?.Line ?? type.Line,
                Members = partMembers
            };
            result.Add(new WrapperFile { Path = PartitionPath(part, source, policy), Namespace = part.Namespace, Category = part.FileCategory, Types = new[] { part } });
        }
        return result;
    }

    private static string ShortHash(string value)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant()[..12];

    private static string PartitionPath(ProjectedType type, string source, ProjectionPolicy policy)
    {
        var root = policy.GeneratedRoot.Trim('/').Replace('\\', '/');
        var product = SafePathPart(type.Product);
        if (!string.IsNullOrWhiteSpace(source))
        {
            var normalized = source.Replace('\\', '/').TrimStart('/');
            if (normalized.EndsWith(policy.FileExtension, StringComparison.OrdinalIgnoreCase)) return string.Join('/', new[] { root, product, normalized });
        }
        return string.Join('/', new[] { root, product, type.FileCategory.ToString(), SafePathPart(type.CSharpName) + policy.FileExtension });
    }

    private static ProjectionTrace TypeTrace(ProjectedType type, DataType data, ContractType? contract, string path)
        => new()
        {
            LogicalId = type.LogicalId,
            Steps = new[]
            {
                Step("source", "enumerate-data-v2-types", data.LogicalId),
                Step("metadata", "optional-wrapper-contract-match", contract?.LogicalId ?? "data-only"),
                Step("namespace", "library-to-netoffice-api-namespace", type.Namespace),
                Step("kind", "typed-data-entity-kind", type.EntityKind.ToString()),
                Step("inheritance", "expand-base-type-graph", string.Join(",", type.EffectiveBaseTypes)),
                Step("capabilities", "derive-from-data-members", string.Join(",", type.Capabilities)),
                Step("runtime", "derive-runtime-requirements", string.Join(",", type.RuntimeRequirements)),
                Step("partition", "current-category-partition", path)
            }
        };

    private static ProjectionTrace MemberTrace(ProjectedMember member, DataMember data, ContractMember? contract, int arity)
        => new()
        {
            LogicalId = member.LogicalId,
            Steps = new[]
            {
                Step("source", "enumerate-data-v2-members", data.LogicalId),
                Step("metadata", "optional-wrapper-contract-match", contract?.Signature ?? "data-only"),
                Step("overload", "expand-trailing-optional-parameters", arity.ToString(CultureInfo.InvariantCulture)),
                Step("signature", "typed-data-signature", member.Signature),
                Step("invocation", "use-data-v2-invocation-evidence", member.Invocation.Operation + ":" + string.Join(",", member.Invocation.ArgumentOrder)),
                Step("conversion", "typed-return-conversion", member.Invocation.ReturnConversion.ToString()),
                Step("docs", "stable-logical-binding-key", member.DocsBindingKey)
            }
        };

    private static ProjectionTraceStep Step(string stage, string rule, string result) => new() { Stage = stage, Rule = rule, Result = result };

    private static List<ProjectedMember> MarkDuplicates(List<ProjectedMember> members)
    {
        var result = new List<ProjectedMember>(members.Count);
        var firstByShape = new Dictionary<string, (int Index, ProjectedMember Member)>(StringComparer.Ordinal);
        foreach (var member in members)
        {
            if (member.Kind is "enum-value" or "field")
            {
                result.Add(member);
                continue;
            }
            var shape = CSharpDeclarationShape(member);
            if (firstByShape.TryGetValue(shape, out var entry) && !string.Equals(entry.Member.DataLogicalId, member.DataLogicalId, StringComparison.Ordinal))
            {
                if (member.Kind is "property" or "indexer" && member.Accessors.Any(accessor => !entry.Member.Accessors.Contains(accessor, StringComparer.Ordinal)))
                {
                    var merged = entry.Member with
                    {
                        Accessors = entry.Member.Accessors.Concat(member.Accessors).Distinct(StringComparer.Ordinal).OrderBy(static accessor => accessor == "get" ? 0 : 1).ToArray(),
                        AccessorInvocations = entry.Member.AccessorInvocations.Concat(member.AccessorInvocations).GroupBy(static plan => plan.OperationKind).Select(static group => group.First()).OrderBy(static plan => plan.OperationKind).ToArray()
                    };
                    result[entry.Index] = merged;
                    firstByShape[shape] = (entry.Index, merged);
                }
                result.Add(member with { DuplicateOf = entry.Member.DataLogicalId });
            }
            else
            {
                firstByShape[shape] = (result.Count, member);
                result.Add(member);
            }
        }
        return result;
    }

    private static string CSharpDeclarationShape(ProjectedMember member)
    {
        var parameters = string.Join(",", member.ParameterList.Select(static parameter => parameter.RefKind + ":" + parameter.Type));
        var owner = member.EmissionOwnerLogicalId + "|";
        return owner + (member.Kind switch
        {
            "indexer" => "indexer|" + parameters,
            "constructor" => "constructor|" + parameters,
            "property" => "property|" + member.CSharpName,
            "event" => "event|" + member.CSharpName,
            _ => member.Kind + "|" + member.CSharpName + "|" + parameters
        });
    }

    private static void ValidateCoverage(DataGraph graph, IReadOnlyList<WrapperFile> files, IReadOnlySet<string> memberIds, IReadOnlySet<string> valueIds, ICollection<ProjectionIssue> issues)
    {
        var projectedTypes = files.SelectMany(static file => file.Types).Select(static type => type.DataLogicalId).ToHashSet(StringComparer.Ordinal);
        foreach (var type in graph.Types.Where(type => !projectedTypes.Contains(type.LogicalId))) issues.Add(new("coverage.type", "Data v2 type was not projected.", type.LogicalId));
        foreach (var member in graph.Members.Where(member => !memberIds.Contains(member.LogicalId))) issues.Add(new("coverage.member", "Data v2 member was not projected.", member.LogicalId));
        foreach (var duplicate in files.GroupBy(static file => file.Path, StringComparer.Ordinal).Where(static group => group.Count() > 1))
            issues.Add(new("coverage.file-conflict", $"Multiple Data v2 types project to {duplicate.Key}.", duplicate.Key));
        foreach (var value in graph.Values.Where(value => !valueIds.Contains(value.LogicalId))) issues.Add(new("coverage.value", "Data v2 value was not projected.", value.LogicalId));
    }

    private static void ValidateContract(WrapperContract contract, string expectedSchema, string? expectedDigest)
    {
        var issues = new List<ProjectionIssue>();
        if (!string.Equals(contract.SchemaVersion, expectedSchema, StringComparison.Ordinal)) issues.Add(new("contract.schema", $"Expected {expectedSchema}, got {contract.SchemaVersion}.", "schemaVersion"));
        if (!string.Equals(contract.ContractKind, "NetOffice.WrapperContract", StringComparison.Ordinal)) issues.Add(new("contract.kind", "Unexpected contract kind.", "contractKind"));
        if (contract.Source is null || string.IsNullOrWhiteSpace(contract.Source.Api)) issues.Add(new("contract.source", "Source.Api is required.", "source.api"));
        if (contract.Types is null) issues.Add(new("contract.types", "Types is required.", "types"));
        else
        {
            foreach (var duplicate in contract.Types.GroupBy(static type => type.LogicalId, StringComparer.Ordinal).Where(static group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1)) issues.Add(new("contract.type-id", "Contract type logical IDs must be unique and non-empty.", duplicate.Key));
            foreach (var type in contract.Types)
            {
                if (string.IsNullOrWhiteSpace(type.LogicalId) || string.IsNullOrWhiteSpace(type.Name) || string.IsNullOrWhiteSpace(type.Signature) || string.IsNullOrWhiteSpace(type.Source) || type.Line < 1) issues.Add(new("contract.type-shape", $"Contract type {type.LogicalId} is incomplete.", type.LogicalId));
                if (type.Members is null) { issues.Add(new("contract.members", "Members is required.", type.LogicalId)); continue; }
                foreach (var member in type.Members)
                    if (string.IsNullOrWhiteSpace(member.Name) || string.IsNullOrWhiteSpace(member.Kind) || string.IsNullOrWhiteSpace(member.Signature) || string.IsNullOrWhiteSpace(member.Source) || member.Line < 1) issues.Add(new("contract.member-shape", $"Contract member {member.Name} is incomplete.", type.LogicalId));
            }
        }
        if (contract.Unknowns is null || contract.Unknowns.Count != 0) issues.Add(new("contract.unknown", "Contract contains unknown records and cannot be projected safely.", "unknowns"));
        if (contract.Ambiguities is null || contract.Ambiguities.Count != 0) issues.Add(new("contract.ambiguity", "Contract contains unresolved ambiguity records.", "ambiguities"));
        var digest = ContractDigest(contract);
        if (expectedDigest is not null && !string.Equals(expectedDigest, digest, StringComparison.OrdinalIgnoreCase)) issues.Add(new("contract.expected-digest", $"Expected pinned digest {expectedDigest}, got {digest}.", "digest"));
        if (issues.Count != 0) throw new ProjectionValidationException(issues);
    }

    public static string ContractDigest(WrapperContract contract)
    {
        var json = System.Text.Json.JsonSerializer.Serialize(contract, ProjectionPolicy.JsonOptions);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant();
    }

    private static int ParameterCount(string? parameters)
    {
        if (string.IsNullOrWhiteSpace(parameters) || parameters.Trim() is "()" or "[]") return 0;
        var value = parameters.Trim();
        if ((value.StartsWith("(", StringComparison.Ordinal) && value.EndsWith(")", StringComparison.Ordinal))
            || (value.StartsWith("[", StringComparison.Ordinal) && value.EndsWith("]", StringComparison.Ordinal))) value = value[1..^1];
        var depth = 0;
        var count = 1;
        var quote = false;
        foreach (var character in value)
        {
            if (character == '"') quote = !quote;
            if (quote) continue;
            if (character is '<' or '(' or '[') depth++;
            else if (character is '>' or ')' or ']') depth--;
            else if (character == ',' && depth == 0) count++;
        }
        return string.IsNullOrWhiteSpace(value) ? 0 : count;
    }

    private static string DocsKey(ProjectionPolicy policy, string id) => policy.DocsKeyPrefix + ":" + id.Replace("/", ".", StringComparison.Ordinal);

    private static string Sanitize(string value, ProjectionNameRules rules)
    {
        var replacement = string.IsNullOrEmpty(rules.InvalidCharacterReplacement) ? "_" : rules.InvalidCharacterReplacement;
        var chars = value.Select(character => char.IsLetterOrDigit(character) || character == '_' ? character : replacement[0]).ToArray();
        var result = new string(chars);
        if (result.Length == 0) result = "_";
        if (char.IsDigit(result[0])) result = "_" + result;
        if ((rules.ReservedWords ?? Array.Empty<string>()).Contains(result, StringComparer.Ordinal)) result = "@" + result;
        return result;
    }

    private static string SafePathPart(string value) => Sanitize(value, new ProjectionNameRules { ReservedWords = Array.Empty<string>() });
    private static string DataLocation(Provenance provenance) => provenance.Location?.Split(':')[0] ?? provenance.SourcePath;
    private static int SourceLine(Provenance provenance)
    {
        var location = provenance.Location;
        if (location is null) return 0;
        var colon = location.LastIndexOf(':');
        return colon >= 0 && int.TryParse(location[(colon + 1)..].Split('/')[0], NumberStyles.None, CultureInfo.InvariantCulture, out var line) ? line : 0;
    }

    private static ProjectedType ApplyTypeOverrides(ProjectedType type, ProjectionPolicy policy, ICollection<ProjectionIssue> issues)
    {
        foreach (var item in policy.Overrides.Where(static item => item.MemberLogicalId is null && item.MemberName is null))
        {
            var matched = MatchesType(item, type);
            if ((matched ? 1 : 0) != item.ExpectedMatches) { issues.Add(new("policy.override.stale", $"Override {item.Id ?? item.Property} matched {(matched ? 1 : 0)}, expected {item.ExpectedMatches}.", item.Id ?? item.Property)); continue; }
            if (!matched) continue;
            type = item.Property switch
            {
                "csharpName" => type with { CSharpName = item.Value },
                "docsBindingKey" => type with { DocsBindingKey = item.Value },
                "namespace" => type with { Namespace = item.Value },
                "runtimeCapability" => type with { RuntimeRequirements = type.RuntimeRequirements.Concat(new[] { item.Value }).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToArray() },
                _ => type
            };
        }
        return type;
    }

    private static void ApplyMemberOverrides(List<WrapperFile> files, ProjectionPolicy policy, ICollection<ProjectionIssue> issues)
    {
        foreach (var item in policy.Overrides.Where(static item => item.MemberLogicalId is not null || item.MemberName is not null))
        {
            var matchCount = files.SelectMany(static file => file.Types).Sum(type => type.Members.Count(member => MatchesMember(item, type, member)));
            if (matchCount != item.ExpectedMatches) issues.Add(new("policy.override.stale", $"Override {item.Id ?? item.Property} matched {matchCount}, expected {item.ExpectedMatches}.", item.Id ?? item.Property));
            for (var fileIndex = 0; fileIndex < files.Count; fileIndex++)
            {
                var file = files[fileIndex];
                files[fileIndex] = file with
                {
                    Types = file.Types.Select(type => type with
                    {
                        Members = type.Members.Select(member => MatchesMember(item, type, member) ? ApplyMemberOverride(member, item) : member).ToArray()
                    }).ToArray()
                };
            }
        }
    }

    private static ProjectedMember ApplyMemberOverride(ProjectedMember member, ProjectionOverride item)
        => item.Property switch
        {
            "csharpName" => member with { CSharpName = item.Value },
            "signature" => member with { Signature = item.Value },
            "docsBindingKey" => member with { DocsBindingKey = item.Value },
            "invocationOperation" => member with { Invocation = member.Invocation with { Operation = item.Value } },
            "runtimeCapability" => member with { RuntimeRequirements = member.RuntimeRequirements.Concat(new[] { item.Value }).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToArray() },
            _ => member
        };

    private static void AppendOverrideTraces(IReadOnlyList<WrapperFile> files, ProjectionPolicy policy, List<ProjectionTrace> traces)
    {
        foreach (var item in policy.Overrides)
        {
            IEnumerable<string> logicalIds;
            if (item.MemberLogicalId is not null || item.MemberName is not null)
                logicalIds = files.SelectMany(static file => file.Types).SelectMany(type => type.Members.Where(member => MatchesMember(item, type, member))).Select(static member => member.LogicalId);
            else
                logicalIds = files.SelectMany(static file => file.Types).Where(type => MatchesType(item, type)).Select(static type => type.LogicalId);
            foreach (var logicalId in logicalIds)
            {
                var index = traces.FindIndex(trace => string.Equals(trace.LogicalId, logicalId, StringComparison.Ordinal));
                if (index >= 0)
                    traces[index] = traces[index] with { Steps = traces[index].Steps.Concat(new[] { Step("policy", "typed-override:" + (item.Id ?? item.Property), item.Property + "=" + item.Value) }).ToArray() };
            }
        }
    }

    private static bool MatchesType(ProjectionOverride item, ProjectedType type)
        => item.TypeLogicalId is not null && (string.Equals(item.TypeLogicalId, type.LogicalId, StringComparison.Ordinal) || string.Equals(item.TypeLogicalId, type.DataLogicalId, StringComparison.Ordinal));

    private static bool MatchesMember(ProjectionOverride item, ProjectedType type, ProjectedMember member)
    {
        if (item.TypeLogicalId is not null && !MatchesType(item, type)) return false;
        if (item.MemberLogicalId is not null && !string.Equals(item.MemberLogicalId, member.LogicalId, StringComparison.Ordinal) && !string.Equals(item.MemberLogicalId, member.DataLogicalId, StringComparison.Ordinal)) return false;
        if (item.MemberName is not null && !string.Equals(item.MemberName, member.Name, StringComparison.Ordinal)) return false;
        return item.MemberLogicalId is not null || item.MemberName is not null;
    }

    private static bool HasCustomEventAccessors(ContractMember? metadata)
        => metadata is not null
           && metadata.Kind.Equals("event", StringComparison.OrdinalIgnoreCase)
           && (metadata.Signature.TrimEnd().EndsWith("{", StringComparison.Ordinal)
               || metadata.InvocationText.Any(static line => line.Contains("+= value", StringComparison.Ordinal)
                                                            || line.Contains("-= value", StringComparison.Ordinal)));

    private static string? ContractEventBackingField(ContractMember? metadata)
    {
        if (metadata is null) return null;
        foreach (var line in metadata.InvocationText)
        {
            var match = EventBackingFieldRegex.Match(line);
            if (match.Success) return match.Groups["field"].Value;
        }
        return null;
    }

    private static ProjectionArgumentPacking ContractArgumentPacking(IEnumerable<string> invocationText)
        => invocationText.Any(static line => line.Contains("new object[]{", StringComparison.Ordinal)
                                             || line.Contains("new object[] {", StringComparison.Ordinal))
            ? ProjectionArgumentPacking.ObjectArray
            : ProjectionArgumentPacking.Flat;

    private static ProjectionObjectArrayStyle ContractObjectArrayStyle(IEnumerable<string> invocationText)
        => invocationText.Any(static line => line.Contains("new object[] {", StringComparison.Ordinal))
            ? ProjectionObjectArrayStyle.Spaced
            : ProjectionObjectArrayStyle.Compact;

    private static ProjectionInvocationApi ContractInvocationApi(IEnumerable<string> invocationText, ProjectionInvocationApi fallback)
    {
        var lines = invocationText as IReadOnlyCollection<string> ?? invocationText.ToArray();
        if (lines.Any(static line => line.Contains("Invoker.", StringComparison.Ordinal))) return ProjectionInvocationApi.Invoker;
        if (lines.Any(static line => line.Contains("Factory.", StringComparison.Ordinal))) return ProjectionInvocationApi.Factory;
        return fallback;
    }

    private static ProjectionInvocationCallKind ContractInvocationCallKind(IEnumerable<string> invocationText)
    {
        var text = string.Join("\n", invocationText);
        if (text.Contains("ExecuteReferencePropertySet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.ReferencePropertySet;
        if (text.Contains("ExecuteVariantPropertySet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.VariantPropertySet;
        if (text.Contains("ExecuteEnumPropertySet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.EnumPropertySet;
        if (text.Contains("ExecuteValuePropertySet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.ValuePropertySet;
        if (text.Contains("ExecutePropertySet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.PropertySet;
        if (text.Contains("MethodGet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.MethodGet;
        if (text.Contains("PropertyGet", StringComparison.Ordinal)) return ProjectionInvocationCallKind.PropertyGet;
        return ProjectionInvocationCallKind.Auto;
    }

    private sealed record ContractMemberMatch(ContractMember? Member, string OwnerLogicalId, ProjectionMemberEmissionDisposition Disposition);

    private sealed record TypeDescriptor(ProjectionEntityKind EntityKind, ProjectionFileCategory Category, string CSharpKind);

    private sealed class VersionComparer : IComparer<string>
    {
        public static readonly VersionComparer Instance = new();
        public int Compare(string? left, string? right)
        {
            if (ReferenceEquals(left, right)) return 0;
            if (left is null) return -1;
            if (right is null) return 1;
            if (decimal.TryParse(left, NumberStyles.Number, CultureInfo.InvariantCulture, out var leftNumber) && decimal.TryParse(right, NumberStyles.Number, CultureInfo.InvariantCulture, out var rightNumber)) return leftNumber.CompareTo(rightNumber);
            return StringComparer.Ordinal.Compare(left, right);
        }
    }
}
