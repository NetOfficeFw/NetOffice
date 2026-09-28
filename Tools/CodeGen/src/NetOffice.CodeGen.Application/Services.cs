using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using NetOffice.CodeGen.Data;
using NetOffice.CodeGen.Docs;
using NetOffice.CodeGen.Emit;
using NetOffice.CodeGen.Projection;
using NetOffice.CodeGen.Storage;
using EmitAccessor = NetOffice.CodeGen.Emit.WrapperAccessor;
using EmitDocumentation = NetOffice.CodeGen.Emit.WrapperDocumentation;
using EmitFile = NetOffice.CodeGen.Emit.WrapperFile;
using EmitMember = NetOffice.CodeGen.Emit.WrapperMember;
using EmitParameter = NetOffice.CodeGen.Emit.WrapperParameter;
using EmitType = NetOffice.CodeGen.Emit.WrapperType;
using ProjectedFile = NetOffice.CodeGen.Projection.WrapperFile;

namespace NetOffice.CodeGen.Application;

public sealed record CommandRequest(
    string Command,
    string? DataPath,
    string? DocsPath,
    string? SourcePath,
    string? OutputPath,
    string? Id,
    string? DocsProfile,
    string? ReportPath,
    bool Locked,
    bool Check,
    bool Apply,
    string? ContractPath = null,
    string? PolicyPath = null,
    string? Projects = null,
    bool NoCache = false,
    string? ExpectedPath = null,
    string? ActualPath = null,
    bool FromData = false,
    bool IsolatedOutput = false);

public sealed record CommandResult(
    int ExitCode,
    string Message,
    IReadOnlyList<string> ChangedPaths,
    CommandReportDetails? Details = null);

public sealed record CopiedPathReport(string Path, string Classification, string Source);

public sealed record CommandReportDetails
{
    public IReadOnlyDictionary<string, string> InputHashes { get; init; } = new Dictionary<string, string>();
    public IReadOnlyDictionary<string, string> ContractFiles { get; init; } = new Dictionary<string, string>();
    public IReadOnlyDictionary<string, long> EntityCounts { get; init; } = new Dictionary<string, long>();
    public IReadOnlyList<string> Products { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> EmittedPaths { get; init; } = Array.Empty<string>();
    public IReadOnlyList<CopiedPathReport> CopiedPaths { get; init; } = Array.Empty<CopiedPathReport>();
    public IReadOnlyList<ProjectionTrace> Traces { get; init; } = Array.Empty<ProjectionTrace>();
    public string? OutputTreeHash { get; init; }
    public bool CacheUsed { get; init; }
    public long ElapsedMilliseconds { get; set; }
    public long PeakManagedBytes { get; set; }
}

public static class ApplicationService
{
    private static readonly string[] GeneratedCategoryNames =
    {
        "Classes", "Constants", "DispatchInterfaces", "Enums", "Events", "Interfaces", "Modules", "Records", "TypeDefs"
    };

    public static CommandResult Execute(CommandRequest request, CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        using var memory = new ManagedMemorySampler();
        CommandResult result;
        try
        {
            result = request.Command switch
            {
                "generate" => Generate(request, cancellationToken),
                "diff" => Diff(request),
                "explain" => Explain(request, cancellationToken),
                "verify" => Fail("verify is not implemented."),
                "bootstrap-ownership" => Fail("bootstrap-ownership is not implemented."),
                "docs" => Fail("docs sync is not implemented."),
                "merge" => Fail("merge is not implemented."),
                "import-typelib" => Fail("import-typelib is not implemented."),
                _ => Fail($"Unknown command '{request.Command}'. Use --help.")
            };
        }
        catch (OperationCanceledException)
        {
            result = new CommandResult(130, "Operation cancelled.", Array.Empty<string>());
        }
        catch (Exception error)
        {
            result = new CommandResult(1, error.Message, Array.Empty<string>());
        }

        stopwatch.Stop();
        if (result.Details is { } details)
        {
            details.ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            details.PeakManagedBytes = memory.PeakBytes;
        }
        WriteReport(request.ReportPath, request, result, stopwatch.ElapsedMilliseconds, memory.PeakBytes);
        return result;
    }

    private static CommandResult Generate(CommandRequest request, CancellationToken cancellationToken)
    {
        ValidateGenerationMode(request);
        var emitter = new CSharpEmitter(new EmitterOptions
        {
            ProductName = "NetOffice",
            GeneratedBy = "NetOffice.CodeGen"
        });
        ProjectionInputs inputs;
        DocumentationResult? documentation;
        EmittedFile[] emitted;
        if (request.DocsProfile?.Equals(nameof(DocsProfile.Baseline), StringComparison.OrdinalIgnoreCase) == true)
        {
            if (string.IsNullOrWhiteSpace(request.ContractPath)) throw new ArgumentException("--contract is required by --docs-profile baseline.");
            var session = new BaselineDocumentationSession(new DocsOptions(
                request.DocsPath ?? Environment.CurrentDirectory,
                Profile: DocsProfile.Baseline,
                Locked: request.Locked,
                BaselineContractPath: request.ContractPath),
                retainDocuments: false);
            var emittedFiles = new global::System.Collections.Concurrent.ConcurrentBag<EmittedFile>();
            inputs = LoadProjectionInputs(request, cancellationToken, consumeProduct: productFiles =>
            {
                var productDocumentation = BindDocumentation(request, productFiles, session);
                Parallel.ForEach(productFiles, new ParallelOptions
                {
                    CancellationToken = cancellationToken,
                    MaxDegreeOfParallelism = Math.Max(1, Math.Min(Environment.ProcessorCount, 8))
                }, file =>
                {
                    var emitInput = ToEmit(file, productDocumentation, request.IsolatedOutput);
                    ValidateComposedEmitFile(emitInput);
                    emittedFiles.Add(emitter.Emit(emitInput));
                });
            });
            documentation = session.Complete();
            emitted = emittedFiles.ToArray();
        }
        else
        {
            inputs = LoadProjectionInputs(request, cancellationToken);
            documentation = BindDocumentation(request, inputs);
            CollectStageGarbage();
            emitted = new EmittedFile[inputs.Files.Count];
            Parallel.For(0, inputs.Files.Count, new ParallelOptions
            {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = Math.Max(1, Math.Min(Environment.ProcessorCount, 8))
            }, index =>
            {
                var emitInput = ToEmit(inputs.Files[index], documentation, request.IsolatedOutput);
                ValidateComposedEmitFile(emitInput);
                emitted[index] = emitter.Emit(emitInput);
            });
        }
        cancellationToken.ThrowIfCancellationRequested();
        Array.Sort(emitted, static (left, right) => StringComparer.Ordinal.Compare(left.RelativePath, right.RelativePath));
        var output = Required(request.OutputPath, "--output");
        var companions = request.IsolatedOutput
            ? CollectCompanions(ExistingDirectory(request.SourcePath, "--source"), output, inputs.Products, request.ContractPath)
            : Array.Empty<CompanionFile>();
        EnsureUniqueDesiredPaths(emitted, companions);

        var previousManifest = ReadManifest(output);
        var plan = GenerationPlanBuilder.Build(output, emitted, previousManifest, new PlanOptions
        {
            CheckOnly = request.Check,
            GeneratorVersion = "application/v3",
            MaxDegreeOfParallelism = Math.Max(1, Math.Min(Environment.ProcessorCount, 8))
        });
        var companionChanges = FindCompanionChanges(output, companions);
        var changed = plan.Operations.Select(static operation => operation.RelativePath)
            .Concat(companionChanges)
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToArray();
        var treeHash = TreeHasher.ComputeFiles(
            emitted.Select(static file => (file.RelativePath, (ReadOnlyMemory<byte>)file.Bytes))
                .Concat(companions.Select(static file => (file.RelativePath, (ReadOnlyMemory<byte>)file.Bytes))));
        var details = CreateDetails(inputs, documentation, emitted, companions, treeHash);

        if (request.Check)
        {
            return changed.Length == 0
                ? new CommandResult(0, "Generation check is clean.", changed, details)
                : new CommandResult(2, $"Generation drift detected in {changed.Length.ToString(CultureInfo.InvariantCulture)} path(s).", changed, details);
        }

        cancellationToken.ThrowIfCancellationRequested();
        plan.Apply();
        WriteCompanions(output, companions, cancellationToken);
        return new CommandResult(
            0,
            request.FromData ? "Exploratory generation completed." : "Generation completed.",
            changed,
            details);
    }

    private static CommandResult Explain(CommandRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("explain requires a logical id.");
        ValidateReadMode(request);
        var inputs = LoadProjectionInputs(request, cancellationToken, includeTraces: true);
        var requestedIds = inputs.Files.SelectMany(static file => file.Types)
            .Where(type => string.Equals(type.DataLogicalId, request.Id, StringComparison.Ordinal) || string.Equals(type.LogicalId, request.Id, StringComparison.Ordinal))
            .Select(static type => type.LogicalId)
            .Append(request.Id)
            .ToHashSet(StringComparer.Ordinal);
        var traces = inputs.Traces.Where(trace => requestedIds.Contains(trace.LogicalId)).ToArray();
        if (traces.Length == 0)
            traces = inputs.Traces.Where(trace => trace.LogicalId.Contains(request.Id, StringComparison.Ordinal)).OrderBy(static trace => trace.LogicalId, StringComparer.Ordinal).ToArray();
        if (traces.Length == 0) return new CommandResult(1, $"No projection trace found for '{request.Id}'.", Array.Empty<string>(), CreateDetails(inputs, null, Array.Empty<EmittedFile>(), Array.Empty<CompanionFile>(), null));
        var message = JsonSerializer.Serialize(traces, ReportJsonOptions);
        var details = CreateDetails(inputs, null, Array.Empty<EmittedFile>(), Array.Empty<CompanionFile>(), null) with { Traces = traces };
        return new CommandResult(0, message, Array.Empty<string>(), details);
    }

    [global::System.Runtime.CompilerServices.MethodImpl(global::System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static ProjectionInputs LoadProjectionInputs(
        CommandRequest request,
        CancellationToken cancellationToken,
        bool includeTraces = false,
        Action<IReadOnlyList<ProjectedFile>>? consumeProduct = null)
    {
        var graphInput = LoadGraph(request.DataPath);
        var policyPath = ExistingFile(request.PolicyPath, "--policy");
        var policy = ProjectionPolicy.Parse(File.ReadAllText(policyPath));
        var products = SelectProducts(graphInput.Graph, request.Projects);
        var contracts = LoadContracts(request.ContractPath, products);
        if (!request.FromData && contracts.Count == 0) throw new ArgumentException("--contract is required unless --from-data/--exploratory is used.");
        if (request.Locked)
        {
            var ignoredProducts = graphInput.Graph.Projects.Where(static project => project.Ignore).Select(static project => project.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var missing = products.Where(product => !ignoredProducts.Contains(product) && !contracts.Any(contract => contract.Api.Equals(product, StringComparison.OrdinalIgnoreCase))).ToArray();
            if (missing.Length != 0) throw new ArgumentException($"Locked generation requires a Wrapper Contract for every selected API product: {string.Join(", ", missing)}.");
        }

        var selectedContracts = contracts
            .Where(contract => products.Contains(contract.Api, StringComparer.OrdinalIgnoreCase))
            .GroupBy(static contract => contract.Api, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                static group => group.Key,
                group => group.Count() == 1 ? group.Single() : throw new InvalidDataException($"Multiple Wrapper Contracts target product '{group.Key}'."),
                StringComparer.OrdinalIgnoreCase);
        var filesByProduct = new IReadOnlyList<ProjectedFile>[products.Count];
        var tracesByProduct = new IReadOnlyList<ProjectionTrace>[products.Count];
        var projectedFileCounts = new long[products.Count];
        var projectedTypeCounts = new long[products.Count];
        var projectedMemberCounts = new long[products.Count];
        const int projectionBatchSize = 2;
        for (var batchStart = 0; batchStart < products.Count; batchStart += projectionBatchSize)
        {
            var batchEnd = Math.Min(products.Count, batchStart + projectionBatchSize);
            Parallel.For(batchStart, batchEnd, new ParallelOptions
        {
            CancellationToken = cancellationToken,
            MaxDegreeOfParallelism = Math.Max(1, Math.Min(Environment.ProcessorCount, projectionBatchSize))
        }, index =>
        {
            var product = products[index];
            var productGraph = FilterGraphForProduct(graphInput.Graph, product);
            selectedContracts.TryGetValue(product, out var contractInput);
            var result = ProjectProduct(productGraph, contractInput, policy, includeTraces);
            cancellationToken.ThrowIfCancellationRequested();
            var productFiles = result.Files
                .Select(file => file with { Types = file.Types
                    .Where(type => type.Product.Equals(product, StringComparison.OrdinalIgnoreCase) && type.EmissionDisposition == ProjectionEmissionDisposition.Emit)
                    .Select(type => TrimProjectionForApplication(type, includeTraces))
                    .ToArray() })
                .Where(static file => file.Types.Count != 0)
                .OrderBy(static file => file.Path, StringComparer.Ordinal).ToArray();
            filesByProduct[index] = productFiles;
            projectedFileCounts[index] = productFiles.Length;
            projectedTypeCounts[index] = productFiles.SelectMany(static file => file.Types).Sum(static type => 1 + type.AuxiliaryTypes.Count);
            projectedMemberCounts[index] = productFiles.SelectMany(static file => file.Types).Sum(static type => type.Members.Count);
            if (includeTraces)
            {
                var logicalIds = productFiles.SelectMany(static file => file.Types)
                    .SelectMany(static type => new[] { type.LogicalId }.Concat(type.Members.Select(static member => member.LogicalId)))
                    .ToHashSet(StringComparer.Ordinal);
                tracesByProduct[index] = result.Explain.Where(trace => logicalIds.Contains(trace.LogicalId)).OrderBy(static trace => trace.LogicalId, StringComparer.Ordinal).ToArray();
            }
            else
            {
                tracesByProduct[index] = Array.Empty<ProjectionTrace>();
            }
            result = null!;
            productGraph = null!;
            });
            if (consumeProduct is not null)
            {
                CollectStageGarbage();
                for (var index = batchStart; index < batchEnd; index++)
                {
                    consumeProduct(filesByProduct[index]);
                    filesByProduct[index] = Array.Empty<ProjectedFile>();
                    CollectStageGarbage();
                }
            }
        }

        var files = filesByProduct.SelectMany(static productFiles => productFiles).OrderBy(static file => file.Path, StringComparer.Ordinal).ToArray();
        var traces = tracesByProduct.SelectMany(static productTraces => productTraces).OrderBy(static trace => trace.LogicalId, StringComparer.Ordinal).ToArray();
        var selectedProjectIds = graphInput.Graph.Projects.Where(project => products.Contains(project.Name, StringComparer.OrdinalIgnoreCase)).Select(static project => project.LogicalId).ToHashSet(StringComparer.Ordinal);
        var selectedTypeIds = graphInput.Graph.Types.Where(type => selectedProjectIds.Contains(type.ProjectId)).Select(static type => type.LogicalId).ToHashSet(StringComparer.Ordinal);
        var counts = new Dictionary<string, long>(StringComparer.Ordinal)
        {
            ["inputProducts"] = graphInput.Graph.Projects.Count,
            ["inputTypes"] = graphInput.Graph.Types.Count,
            ["inputMembers"] = graphInput.Graph.Members.Count,
            ["inputValues"] = graphInput.Graph.Values.Count,
            ["selectedProducts"] = products.Count,
            ["selectedTypes"] = selectedTypeIds.Count,
            ["selectedMembers"] = graphInput.Graph.Members.Count(member => selectedTypeIds.Contains(member.TypeId)),
            ["selectedValues"] = graphInput.Graph.Values.Count(value => selectedTypeIds.Contains(value.TypeId)),
            ["projectedFiles"] = projectedFileCounts.Sum(),
            ["projectedTypes"] = projectedTypeCounts.Sum(),
            ["projectedMembersAndValues"] = projectedMemberCounts.Sum()
        };
        var contractHashes = contracts.OrderBy(static contract => contract.RelativePath, StringComparer.Ordinal)
            .ToDictionary(static contract => contract.RelativePath, static contract => contract.Sha256, StringComparer.Ordinal);
        var aggregateContractHash = contractHashes.Count == 0
            ? HashBytes(Array.Empty<byte>())
            : TreeHasher.ComputeFiles(contracts.Select(static contract => (contract.RelativePath, (ReadOnlyMemory<byte>)contract.Bytes)));
        return new ProjectionInputs(
            graphInput.Graph.Digest,
            graphInput.FileHash,
            policy.Digest,
            aggregateContractHash,
            contractHashes,
            products,
            files,
            traces,
            counts);
    }

    private static ProjectionResult ProjectProduct(DataGraph graph, ContractInput? input, ProjectionPolicy policy, bool includeTraces)
    {
        var contract = input is null ? SyntheticContract() : WrapperContract.Parse(Encoding.UTF8.GetString(input.Bytes));
        return ProjectionEngine.Project(graph, contract, policy, new ProjectionOptions { IncludeTraces = includeTraces });
    }


    private static DocumentationResult? BindDocumentation(CommandRequest request, ProjectionInputs inputs)
        => BindDocumentation(request, inputs.Files, null);

    private static DocumentationResult? BindDocumentation(
        CommandRequest request,
        IReadOnlyList<ProjectedFile> files,
        BaselineDocumentationSession? baselineSession)
    {
        if (string.IsNullOrWhiteSpace(request.DocsProfile)) return null;
        if (!Enum.TryParse<DocsProfile>(request.DocsProfile, true, out var profile)) throw new ArgumentException($"Unknown documentation profile '{request.DocsProfile}'.");
        if (profile == DocsProfile.Baseline && string.IsNullOrWhiteSpace(request.ContractPath)) throw new ArgumentException("--contract is required by --docs-profile baseline.");
        var targets = new List<DocumentationTarget>();
        var projectedTypes = files.SelectMany(static file => file.Types)
            .OrderBy(static type => type.LogicalId, StringComparer.Ordinal)
            .ThenBy(static type => type.ContractPartSource ?? type.Source, StringComparer.Ordinal)
            .ToArray();
        var multipartTypeIds = projectedTypes
            .GroupBy(static type => string.IsNullOrWhiteSpace(type.CanonicalLogicalId) ? type.LogicalId : type.CanonicalLogicalId, StringComparer.Ordinal)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .ToHashSet(StringComparer.Ordinal);
        foreach (var type in projectedTypes)
        {
            var canonicalTypeId = !string.IsNullOrWhiteSpace(type.ContractPartLogicalId)
                ? type.ContractPartLogicalId
                : string.IsNullOrWhiteSpace(type.CanonicalLogicalId) ? type.LogicalId : type.CanonicalLogicalId;
            var multipartTypeId = string.IsNullOrWhiteSpace(type.CanonicalLogicalId) ? type.LogicalId : type.CanonicalLogicalId;
            targets.Add(DocumentationTarget.ForType(
                multipartTypeIds.Contains(multipartTypeId)
                    ? type.LogicalId + "/part/" + HashBytes(Encoding.UTF8.GetBytes(type.ContractPartSource ?? type.Source)).Substring(0, 16)
                    : type.LogicalId,
                ProjectedTypeDocumentationBindingKey(type),
                canonicalTypeId,
                type.CSharpName,
                type.Namespace,
                type.ContractPartSource ?? type.Source));
            foreach (var auxiliary in type.AuxiliaryTypes.OrderBy(static auxiliary => auxiliary.LogicalId, StringComparer.Ordinal))
            {
                targets.Add(DocumentationTarget.ForType(
                    auxiliary.LogicalId,
                    AuxiliaryDocumentationBindingKey(auxiliary),
                    auxiliary.LogicalId,
                    auxiliary.Name,
                    auxiliary.Namespace,
                    auxiliary.Source));
                foreach (var constructor in auxiliary.Constructors.Where(constructor =>
                             !auxiliary.Name.EndsWith("_SinkHelper", StringComparison.Ordinal)
                             || !string.IsNullOrWhiteSpace(constructor.LogicalId)
                             || !string.IsNullOrWhiteSpace(constructor.DocsBindingKey))
                         .OrderBy(static constructor => constructor.LogicalId, StringComparer.Ordinal))
                {
                    var docsBindingKey = ConstructorDocumentationBindingKey(auxiliary.LogicalId, constructor);
                    targets.Add(DocumentationTarget.ForMember(
                        string.IsNullOrWhiteSpace(constructor.LogicalId) ? docsBindingKey : constructor.LogicalId,
                        docsBindingKey,
                        auxiliary.LogicalId,
                        auxiliary.Name,
                        "constructor",
                        constructor.Signature,
                        constructor.Parameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName),
                        auxiliary.Name,
                        auxiliary.Namespace,
                        string.IsNullOrWhiteSpace(constructor.Source) ? auxiliary.Source : constructor.Source));
                }
            }
            foreach (var member in type.Members.OrderBy(static member => member.LogicalId, StringComparer.Ordinal))
            {
                var emissionOwnerId = string.IsNullOrWhiteSpace(member.EmissionOwnerLogicalId) ? canonicalTypeId : member.EmissionOwnerLogicalId;
                var auxiliaryOwner = type.AuxiliaryTypes.FirstOrDefault(auxiliary => auxiliary.LogicalId.Equals(emissionOwnerId, StringComparison.Ordinal));
                targets.Add(DocumentationTarget.ForMember(
                    member.LogicalId,
                    DocumentationBindingKey(member),
                    emissionOwnerId,
                    member.CSharpName,
                    member.Kind,
                    member.Signature,
                    member.ParameterList.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName),
                    auxiliaryOwner?.Name ?? type.CSharpName,
                    auxiliaryOwner?.Namespace ?? type.Namespace,
                    member.Source));
            }
            var projectedMemberIds = type.Members.Select(static member => member.LogicalId).ToHashSet(StringComparer.Ordinal);
            foreach (var member in type.ExactContractMembers.Where(member => !projectedMemberIds.Contains(member.LogicalId)).OrderBy(static member => member.LogicalId, StringComparer.Ordinal))
            {
                targets.Add(DocumentationTarget.ForMember(
                    member.LogicalId,
                    member.DocsBindingKey,
                    canonicalTypeId,
                    member.Name,
                    member.Kind,
                    member.Signature,
                    member.ParameterList.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName),
                    type.CSharpName,
                    type.Namespace,
                    member.Source));
            }
            foreach (var constructor in type.Constructors.Where(static constructor => constructor.IsContractOverlay).OrderBy(static constructor => constructor.LogicalId, StringComparer.Ordinal))
            {
                targets.Add(DocumentationTarget.ForMember(
                    constructor.LogicalId,
                    constructor.DocsBindingKey,
                    canonicalTypeId,
                    type.CSharpName,
                    "constructor",
                    constructor.Signature,
                    constructor.Parameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.CSharpName),
                    type.CSharpName,
                    type.Namespace,
                    constructor.Source));
            }
        }
        var options = new DocsOptions(
            request.DocsPath ?? Environment.CurrentDirectory,
            Profile: profile,
            Locked: request.Locked,
            BaselineContractPath: request.ContractPath);
        if (baselineSession is null) return DocumentationSync.Sync(targets, options);
        var products = files.SelectMany(static file => file.Types).Select(static type => type.Product).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (products.Length != 1) throw new InvalidOperationException("A baseline documentation batch must contain exactly one product.");
        var contractPath = Required(request.ContractPath, "--contract");
        var productContractPath = Directory.Exists(contractPath)
            ? Path.Combine(contractPath, products[0] + ".wrapper-contract.json")
            : contractPath;
        return baselineSession.BindBatch(targets, productContractPath);
    }

    private static void ValidateComposedEmitFile(EmitFile file)
    {
        foreach (var type in file.Types)
        foreach (var member in type.Members)
        foreach (var accessor in member.Accessors.Where(static accessor => accessor.Kind.Equals("set", StringComparison.OrdinalIgnoreCase)))
        {
            if (accessor.Invocation is { } invocation
                && invocation.Kind is WrapperInvocationKind.PropertySet
                    or WrapperInvocationKind.PropertySetValue
                    or WrapperInvocationKind.PropertySetVariant
                    or WrapperInvocationKind.PropertySetEnum
                    or WrapperInvocationKind.PropertyPutRef)
                continue;
            throw new InvalidDataException(
                $"Composed set accessor '{type.Name}.{member.Name}' in '{file.RelativePath}' has invocation kind '{accessor.Invocation?.Kind.ToString() ?? "<none>"}'.");
        }
    }

    private static EmitFile ToEmit(ProjectedFile file, DocumentationResult? documentation, bool isolated)
    {
        var projectedType = file.Types.Single();
        var partition = ExactPartition(file.Path, projectedType.Product, projectedType.FileCategory, projectedType.CSharpName);
        var relativePath = isolated
            ? $"Source/{SafePathPart(projectedType.Product)}/Generated/{partition}"
            : $"{SafePathPart(projectedType.Product)}/Generated/{partition}";
        return new EmitFile
        {
            RelativePath = relativePath,
            Namespace = file.Namespace,
            Usings = new[] { "System", "System.Collections", "System.Collections.Generic", "System.ComponentModel", "System.Runtime.InteropServices", "System.Runtime.InteropServices.ComTypes", "NetOffice", "NetOffice.Attributes", "NetOffice.CollectionsGeneric", "NetRuntimeSystem = System", "ComEventInterface = NetOffice.Attributes.ComEventInterfaceAttribute" },
            Types = file.Types.SelectMany(type => new[] { ToEmitType(type, documentation) }
                .Concat(type.IsPrimaryPart
                    ? type.AuxiliaryTypes
                        .Where(auxiliary => !string.Equals(auxiliary.Name, type.EventSink?.Name ?? type.CSharpName + "_SinkHelper", StringComparison.Ordinal))
                        .Select(auxiliary => ToEmitAuxiliaryType(type, auxiliary, documentation))
                    : Array.Empty<EmitType>())).ToArray()
        };
    }
    private static string ExactPartition(string path, string product, ProjectionFileCategory category, string typeName)
    {
        var segments = path.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);
        var productIndex = Array.FindIndex(segments, segment => segment.Equals(product, StringComparison.OrdinalIgnoreCase));
        return productIndex >= 0 && productIndex + 1 < segments.Length
            ? string.Join("/", segments.Skip(productIndex + 1))
            : $"{category}/{Path.GetFileName(path.Length == 0 ? typeName + ".cs" : path)}";
    }


    private static ProjectedType TrimProjectionForApplication(ProjectedType type, bool preserveExplainIdentity)
    {
        var localForwardNames = type.Members
            .Where(static member => member.Invocation.OperationKind == ProjectionInvocationOperation.LocalForward)
            .Select(static member => member.CSharpName)
            .ToHashSet(StringComparer.Ordinal);
        var exactMembers = type.ExactContractMembers.Where(member =>
            localForwardNames.Contains(member.Name)
            || member.Name.EndsWith(".GetEnumerator", StringComparison.Ordinal)
            || member.Name is "InstanceType" or "LateBindingApiWrapperType"
                or "CreateEventBridge" or "EventBridgeInitialized" or "HasEventRecipients"
                or "GetEventRecipients" or "GetCountOfEventRecipients" or "RaiseCustomEvent" or "DisposeEventBridge"
                or "AssemblyName" or "AssemblyNamespace" or "ComponentGuid" or "AssemblyAttribute" or "Assembly"
                or "Dependencies" or "Contains" or "GetComObjectEnumerator" or "FetchVariantComObjectEnumerator"
                or "GetEnumerator")
            .ToArray();
        var members = type.Members.Select(member => member with
        {
            DataLogicalId = preserveExplainIdentity ? member.DataLogicalId : string.Empty,
            OverloadGroup = string.Empty,
            SupportVersions = Array.Empty<VersionSupport>(),
            Capabilities = Array.Empty<string>(),
            RuntimeRequirements = Array.Empty<string>()
        }).ToArray();
        return type with
        {
            DataLogicalId = preserveExplainIdentity ? type.DataLogicalId : string.Empty,
            LibraryLogicalId = string.Empty,
            DuplicateGroups = Array.Empty<string>(),
            SupportVersions = Array.Empty<VersionSupport>(),
            Capabilities = Array.Empty<string>(),
            Signature = string.Empty,
            ExactContractMembers = exactMembers,
            Members = members
        };
    }

    private static EmitType ToEmitType(ProjectedType type, DocumentationResult? documentation)
    {
        var kind = NormalizeTypeKind(type.Kind, type.EntityKind);
        var modifiers = type.Modifiers.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var runtime = type.IsPrimaryPart && kind == "class" && type.RuntimeRequirements.Contains("netoffice-core", StringComparer.Ordinal)
            ? new WrapperRuntime
            {
                EmitStandardConstructors = false,
                ConstructorProfile = type.EntityKind == ProjectionEntityKind.CoClass ? WrapperConstructorProfile.CoClass : WrapperConstructorProfile.Wrapper,
                ProgId = type.ProgId,
                InstanceTypeContract = RuntimeMemberContract(type, "InstanceType", documentation),
                LateBindingApiWrapperTypeContract = RuntimeMemberContract(type, "LateBindingApiWrapperType", documentation)
            }
            : null;
        var members = ToEmitMembers(type, type.LogicalId, documentation).ToList();
        if (runtime is not null)
        {
            members.AddRange(type.Constructors.Select(constructor => ToEmitConstructor(type, constructor, documentation)));
        }
        var bases = new[] { type.BaseType }.Where(static value => !string.IsNullOrWhiteSpace(value)).Cast<string>()
            .Concat(type.Interfaces).Distinct(StringComparer.Ordinal).ToArray();
        if (type.IsPrimaryPart)
        {
            var providerInterfaces = bases.Where(static baseType => baseType.Contains("IEnumerableProvider<", StringComparison.Ordinal)).ToArray();
            foreach (var collectionInterface in providerInterfaces)
                members.AddRange(CollectionProviderMembers(type, collectionInterface, documentation));
            if (providerInterfaces.Length == 0)
                foreach (var collectionInterface in bases.Where(static baseType =>
                             baseType.Contains("IEnumerable<", StringComparison.Ordinal)
                             && !baseType.Contains("IEnumerableProvider<", StringComparison.Ordinal)))
                    members.AddRange(CollectionEnumeratorMembers(type, collectionInterface, documentation));
        }
        return new EmitType
        {
            Name = type.CSharpName,
            Kind = kind,
            Accessibility = type.Accessibility,
            Partial = modifiers.Contains("partial"),
            Sealed = modifiers.Contains("sealed"),
            Abstract = modifiers.Contains("abstract"),
            Static = modifiers.Contains("static"),
            Runtime = runtime,
            ModuleRuntime = kind == "module"
                ? new WrapperModuleRuntime { InstanceType = "ICOMObject", CoreType = "Core", InvokerType = "Invoker" }
                : null,
            EventBinding = type.IsPrimaryPart ? ToEventBinding(type, documentation) : null,
            ProjectInfo = type.IsPrimaryPart ? ToProjectInfo(type, documentation) : null,
            BaseTypes = bases.ToArray(),
            Attributes = (type.IsPrimaryPart
                ? TypeAttributes(type)
                : TypeAttributes(type).Where(static attribute => attribute.StartsWith("SyntaxBypass", StringComparison.Ordinal))).ToArray(),
            Documentation = FindDocumentation(documentation, ProjectedTypeDocumentationBindingKey(type)),
            Members = members
        };
    }

    private static EmitType ToEmitAuxiliaryType(ProjectedType owner, ProjectedAuxiliaryType auxiliary, DocumentationResult? documentation)
    {
        var modifiers = auxiliary.Modifiers.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var members = ToEmitMembers(owner, auxiliary.LogicalId, documentation).ToList();
        members.AddRange(auxiliary.Constructors.Select(constructor => ToEmitConstructor(
            auxiliary.Name,
            constructor with { DocsBindingKey = ConstructorDocumentationBindingKey(auxiliary.LogicalId, constructor) },
            documentation)));
        var bases = new[] { auxiliary.BaseType }.Where(static value => !string.IsNullOrWhiteSpace(value))
            .Concat(auxiliary.Interfaces).Cast<string>().Distinct(StringComparer.Ordinal).ToArray();
        return new EmitType
        {
            Name = auxiliary.Name,
            Kind = auxiliary.Kind.Trim().ToLowerInvariant(),
            Accessibility = auxiliary.Accessibility,
            Partial = modifiers.Contains("partial"),
            Sealed = modifiers.Contains("sealed"),
            Abstract = modifiers.Contains("abstract"),
            Static = modifiers.Contains("static"),
            BaseTypes = bases,
            Attributes = NormalizeAttributes(auxiliary.Attributes),
            DelegateReturnType = ProjectedManagedType(null, auxiliary.DelegateReturnType ?? "void"),
            DelegateParameters = auxiliary.DelegateParameters.OrderBy(static parameter => parameter.Position).Select(ToEmitParameter).ToArray(),
            Documentation = FindDocumentation(documentation, AuxiliaryDocumentationBindingKey(auxiliary)),
            Members = members
        };
    }

    private static string ProjectedTypeDocumentationBindingKey(ProjectedType type)
        => type.DocsBindingKey + "/source/"
            + HashBytes(Encoding.UTF8.GetBytes(type.ContractPartSource ?? type.Source)).Substring(0, 16);

    private static string AuxiliaryDocumentationBindingKey(ProjectedAuxiliaryType auxiliary)
        => "auxiliary-type:" + auxiliary.LogicalId;

    private static string ConstructorDocumentationBindingKey(string ownerLogicalId, ConstructorPlan constructor)
        => string.IsNullOrWhiteSpace(constructor.DocsBindingKey)
            ? "constructor:" + ownerLogicalId + ":" + constructor.Kind + ":"
                + string.Join(",", constructor.Parameters.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.Type + " " + parameter.CSharpName))
                + ":" + constructor.BaseCall
            : constructor.DocsBindingKey;

    private static WrapperRuntimeMemberContract? RuntimeMemberContract(ProjectedType type, string name, DocumentationResult? documentation, int? parameterCount = null)
        => RuntimeMemberContract(type.ExactContractMembers.FirstOrDefault(member =>
            member.Name.Equals(name, StringComparison.Ordinal)
            && (!parameterCount.HasValue || member.ParameterList.Count == parameterCount.Value)), documentation);

    private static WrapperRuntimeMemberContract? RuntimeMemberContract(ProjectedContractMember? exact, DocumentationResult? documentation)
        => exact is null
            ? null
            : new WrapperRuntimeMemberContract
            {
                Attributes = NormalizeAttributes(exact.Attributes),
                Documentation = FindDocumentation(documentation, exact.DocsBindingKey)
            };

    private static WrapperRuntimeMemberContract? RuntimeMemberContract(ConstructorPlan? exact, DocumentationResult? documentation)
        => exact is null
            ? null
            : new WrapperRuntimeMemberContract
            {
                Attributes = NormalizeAttributes(exact.Attributes),
                Documentation = FindDocumentation(documentation, exact.DocsBindingKey)
            };

    private static WrapperEventBinding? ToEventBinding(ProjectedType type, DocumentationResult? documentation)
    {
        if (type.EventBindings.Count == 0) return null;
        return new WrapperEventBinding
        {
            Sinks = type.EventBindings.Select(static binding => new WrapperEventSinkBinding
            {
                SinkHelperType = binding.SinkHelperType,
                FieldName = binding.FieldName
            }).ToArray(),
            Contracts = new WrapperEventBindingMemberContracts
            {
                CreateEventBridge = RuntimeMemberContract(type, "CreateEventBridge", documentation, 0),
                EventBridgeInitialized = RuntimeMemberContract(type, "EventBridgeInitialized", documentation, 0),
                HasEventRecipients = RuntimeMemberContract(type, "HasEventRecipients", documentation, 0),
                HasNamedEventRecipients = RuntimeMemberContract(type, "HasEventRecipients", documentation, 1),
                GetEventRecipients = RuntimeMemberContract(type, "GetEventRecipients", documentation, 1),
                GetCountOfEventRecipients = RuntimeMemberContract(type, "GetCountOfEventRecipients", documentation, 1),
                RaiseCustomEvent = RuntimeMemberContract(type, "RaiseCustomEvent", documentation, 2),
                DisposeEventBridge = RuntimeMemberContract(type, "DisposeEventBridge", documentation, 0)
            }
        };
    }

    private static WrapperProjectInfo? ToProjectInfo(ProjectedType type, DocumentationResult? documentation)
    {
        if (type.ProjectInfo is null) return null;
        var contains = type.ExactContractMembers.Where(static member => member.Name.Equals("Contains", StringComparison.Ordinal)).ToArray();
        return new WrapperProjectInfo
        {
            AssemblyNamespace = type.ProjectInfo.AssemblyNamespace,
            ComponentGuids = type.ProjectInfo.ComponentGuids.ToArray(),
            Dependencies = type.ProjectInfo.Dependencies.ToArray(),
            Contracts = new WrapperProjectInfoMemberContracts
            {
                Constructor = RuntimeMemberContract(type.Constructors.FirstOrDefault(static constructor => constructor.IsContractOverlay), documentation),
                AssemblyName = RuntimeMemberContract(type, "AssemblyName", documentation, 0),
                AssemblyNamespace = RuntimeMemberContract(type, "AssemblyNamespace", documentation, 0),
                ComponentGuid = RuntimeMemberContract(type, "ComponentGuid", documentation, 0),
                AssemblyAttribute = RuntimeMemberContract(type, "AssemblyAttribute", documentation, 0),
                Assembly = RuntimeMemberContract(type, "Assembly", documentation, 0),
                Dependencies = RuntimeMemberContract(type, "Dependencies", documentation, 0),
                ContainsType = RuntimeMemberContract(contains.FirstOrDefault(member => member.ParameterList.Count == 1 && !member.ParameterList[0].Type.Equals("string", StringComparison.OrdinalIgnoreCase)), documentation),
                ContainsClassName = RuntimeMemberContract(contains.FirstOrDefault(member => member.ParameterList.Count == 1 && member.ParameterList[0].Type.Equals("string", StringComparison.OrdinalIgnoreCase)), documentation)
            }
        };
    }

    private static IReadOnlyList<string> TypeAttributes(ProjectedType type)
    {
        var attributes = type.Attributes.ToList();
        if (type.EventSink is null) return NormalizeAttributes(attributes);
        if (!attributes.Any(static attribute => attribute.Contains("InternalEntityKind.ComEventInterface", StringComparison.Ordinal)))
            attributes.Add("InternalEntity(InternalEntityKind.ComEventInterface)");
        if (!string.IsNullOrWhiteSpace(type.EventSink.InterfaceId) && !attributes.Any(static attribute => attribute.Contains("Guid(", StringComparison.Ordinal)))
            attributes.Add($"Guid(\"{type.EventSink.InterfaceId}\")");
        if (!type.EventSink.Name.Equals(type.EventSink.InterfaceName + "_SinkHelper", StringComparison.Ordinal))
            throw new InvalidDataException($"Event sink '{type.EventSink.Name}' does not match the emitter convention for '{type.EventSink.InterfaceName}'.");
        return NormalizeAttributes(attributes);
    }


    private static string ExactExplicitInterface(ProjectedContractMember? exact, string fallback)
    {
        if (exact is null || string.IsNullOrWhiteSpace(exact.Signature)) return fallback;
        var marker = "." + exact.Name + "(";
        var markerIndex = exact.Signature.IndexOf(marker, StringComparison.Ordinal);
        if (markerIndex < 0) return fallback;
        var prefix = exact.Signature.Substring(0, markerIndex).TrimEnd();
        var separator = prefix.LastIndexOf(' ');
        return separator < 0 ? fallback : prefix.Substring(separator + 1);
    }

    private static IEnumerable<EmitMember> CollectionProviderMembers(ProjectedType type, string collectionInterface, DocumentationResult? documentation)
    {
        ProjectedContractMember? Exact(string name)
            => type.ExactContractMembers.FirstOrDefault(member => member.Name.Equals(name, StringComparison.Ordinal));
        var comObjectExact = Exact("GetComObjectEnumerator");
        var variantExact = Exact("FetchVariantComObjectEnumerator");
        var genericEnumerator = type.ExactContractMembers.FirstOrDefault(static member => member.Name.Equals("GetEnumerator", StringComparison.Ordinal));
        var nativeEnumerator = type.ExactContractMembers.FirstOrDefault(static member => member.Name.EndsWith(".GetEnumerator", StringComparison.Ordinal));

        var comObject = new EmitMember
        {
            Name = "GetComObjectEnumerator",
            Kind = "method",
            Type = "ICOMObject",
            ExplicitInterface = ExactExplicitInterface(comObjectExact, collectionInterface),
            Parameters = new[] { new EmitParameter { Name = "parent", Type = "ICOMObject" } },
            Body = "return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);"
        };
        yield return ApplyExactContractMember(comObject, comObjectExact, documentation);

        var variant = new EmitMember
        {
            Name = "FetchVariantComObjectEnumerator",
            Kind = "method",
            Type = "IEnumerable",
            ExplicitInterface = ExactExplicitInterface(variantExact, collectionInterface),
            Parameters = new[]
            {
                new EmitParameter { Name = "parent", Type = "ICOMObject" },
                new EmitParameter { Name = "enumerator", Type = "ICOMObject" }
            },
            Body = "return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);"
        };
        yield return ApplyExactContractMember(variant, variantExact, documentation);

        var open = collectionInterface.IndexOf('<');
        var close = collectionInterface.LastIndexOf('>');
        var itemType = open >= 0 && close > open ? collectionInterface.Substring(open + 1, close - open - 1) : "object";
        var generic = new EmitMember
        {
            Name = "GetEnumerator",
            Kind = "method",
            Type = genericEnumerator?.ReturnType ?? $"IEnumerator<{itemType}>",
            Body = $"global::System.Collections.IEnumerable innerEnumerator = (this as global::System.Collections.IEnumerable);\nforeach ({itemType} item in innerEnumerator)\n    yield return item;"
        };
        yield return ApplyExactContractMember(generic, genericEnumerator, documentation);

        var native = new EmitMember
        {
            Name = "GetEnumerator",
            Kind = "method",
            Type = nativeEnumerator?.ReturnType ?? "IEnumerator",
            ExplicitInterface = nativeEnumerator is null
                ? "global::System.Collections.IEnumerable"
                : nativeEnumerator.Name.Substring(0, nativeEnumerator.Name.Length - ".GetEnumerator".Length),
            Body = "return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);"
        };
        yield return ApplyExactContractMember(native, nativeEnumerator, documentation);
    }

    private static IEnumerable<EmitMember> CollectionEnumeratorMembers(ProjectedType type, string collectionInterface, DocumentationResult? documentation)
    {
        var genericEnumerator = type.ExactContractMembers.FirstOrDefault(static member => member.Name.Equals("GetEnumerator", StringComparison.Ordinal));
        var nativeEnumerator = type.ExactContractMembers.FirstOrDefault(static member => member.Name.EndsWith(".GetEnumerator", StringComparison.Ordinal));
        var open = collectionInterface.IndexOf('<');
        var close = collectionInterface.LastIndexOf('>');
        var itemType = open >= 0 && close > open ? collectionInterface.Substring(open + 1, close - open - 1) : "object";

        var generic = new EmitMember
        {
            Name = "GetEnumerator",
            Kind = "method",
            Type = genericEnumerator?.ReturnType ?? $"IEnumerator<{itemType}>",
            Body = $"global::System.Collections.IEnumerable innerEnumerator = (this as global::System.Collections.IEnumerable);\nforeach ({itemType} item in innerEnumerator)\n    yield return item;"
        };
        yield return ApplyExactContractMember(generic, genericEnumerator, documentation);

        var enumeratorAsProperty = type.Attributes.Any(static attribute =>
            attribute.Contains("EnumeratorInvoke.Property", StringComparison.Ordinal));
        var native = new EmitMember
        {
            Name = "GetEnumerator",
            Kind = "method",
            Type = nativeEnumerator?.ReturnType ?? "IEnumerator",
            ExplicitInterface = nativeEnumerator is null
                ? "global::System.Collections.IEnumerable"
                : nativeEnumerator.Name.Substring(0, nativeEnumerator.Name.Length - ".GetEnumerator".Length),
            Body = enumeratorAsProperty
                ? "return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);"
                : "return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);"
        };
        yield return ApplyExactContractMember(native, nativeEnumerator, documentation);
    }

    private static EmitMember ApplyExactContractMember(EmitMember member, ProjectedContractMember? exact, DocumentationResult? documentation)
    {
        if (exact is null) return member;
        member.Accessibility = exact.Accessibility;
        member.Attributes = NormalizeAttributes(exact.Attributes);
        member.Parameters = exact.ParameterList.OrderBy(static parameter => parameter.Position).Select(ToEmitParameter).ToArray();
        member.PreserveParameterAttributes = exact.ParameterList.Any(static parameter => parameter.PreserveContractAttributes);
        member.Documentation = FindDocumentation(documentation, exact.DocsBindingKey);
        return member;
    }

    private static IEnumerable<EmitMember> ToEmitMembers(ProjectedType type, string ownerLogicalId, DocumentationResult? documentation)
    {
        var distinctMembers = type.Members.Where(member => member.DuplicateOf is null
            && member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit
            && string.Equals(EffectiveEmissionOwner(type, member), ownerLogicalId, StringComparison.Ordinal))
            .GroupBy(ProjectionMemberSignatureKey, StringComparer.Ordinal)
            .Select(static group => group
                .OrderByDescending(static member => member.ParameterList.Count(static parameter => parameter.SinkArgument is not null))
                .ThenByDescending(static member => member.Attributes.Count)
                .ThenBy(static member => member.EventSinkOnly)
                .ThenBy(static member => member.IsContractDerived)
                .ThenBy(static member => member.LogicalId, StringComparer.Ordinal)
                .First())
            .ToArray();
        var allAccessorMembers = distinctMembers.Where(static member => member.Kind is "property" or "indexer" && !string.IsNullOrWhiteSpace(member.AccessorGroupId)).ToArray();
        var indexedGroups = allAccessorMembers.Where(static member => member.Kind == "indexer").Select(static member => member.AccessorGroupId!).ToHashSet(StringComparer.Ordinal);
        var accessorMembers = allAccessorMembers.Where(member => !(member.Kind == "property" && member.ParameterList.Count == 0 && indexedGroups.Contains(member.AccessorGroupId!))).ToArray();
        var accessorIds = allAccessorMembers.Select(static member => member.LogicalId).ToHashSet(StringComparer.Ordinal);
        foreach (var member in distinctMembers.Where(member => !accessorIds.Contains(member.LogicalId)).OrderBy(static member => member.LogicalId, StringComparer.Ordinal))
        {
            yield return ToEmitMember(type, ApplyLocalForwardContract(type, member), documentation, type.EntityKind == ProjectionEntityKind.Constants);
        }
        foreach (var group in accessorMembers.GroupBy(AccessorProjectionKey, StringComparer.Ordinal).OrderBy(static group => group.Key, StringComparer.Ordinal))
        {
            var direct = group.ToArray();
            var plannedAccessors = direct.SelectMany(static member =>
                member.AccessorInvocations.Select(invocation => member with { Invocation = invocation })).ToArray();
            var plannedAccessorKinds = plannedAccessors.Select(static member => EmitAccessorKind(member.Invocation)).ToHashSet(StringComparer.Ordinal);
            var projected = direct.Where(member => !plannedAccessorKinds.Contains(EmitAccessorKind(member.Invocation)))
                .Concat(plannedAccessors.GroupBy(static member => EmitAccessorKind(member.Invocation), StringComparer.Ordinal).Select(static candidates => candidates.First()))
                .OrderBy(static member => AccessorOrder(member.Invocation))
                .ThenBy(static member => member.LogicalId, StringComparer.Ordinal)
                .ToArray();
            var getter = projected.FirstOrDefault(static member => EmitAccessorKind(member.Invocation) == "get");
            var representative = getter ?? projected[0];
            var modifiers = representative.Modifiers.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var accessors = projected.GroupBy(static member => EmitAccessorKind(member.Invocation), StringComparer.Ordinal)
                .Select(static candidates => candidates.First())
                .OrderBy(static member => AccessorOrder(member.Invocation))
                .Select(member => new EmitAccessor
                {
                    Kind = EmitAccessorKind(member.Invocation),
                    Invocation = MapInvocation(member, type.Namespace),
                    Attributes = Array.Empty<string>()
                }).ToArray();
            yield return new EmitMember
            {
                Name = representative.CSharpName,
                Kind = representative.Kind,
                Accessibility = representative.Accessibility,
                Type = ProjectedManagedType(getter is not null && getter.Invocation.HasContractInvocation ? null : getter?.Invocation.ResultTypeReference ?? representative.Invocation.ResultTypeReference, getter?.ReturnType ?? representative.ReturnType ?? "object"),
                Static = modifiers.Contains("static"),
                Virtual = modifiers.Contains("virtual"),
                Override = modifiers.Contains("override"),
                Abstract = modifiers.Contains("abstract"),
                New = modifiers.Contains("new"),
                Attributes = NormalizeAttributes(projected.SelectMany(static member => member.Attributes).Distinct(StringComparer.Ordinal)),
                Parameters = representative.ParameterList.OrderBy(static parameter => parameter.Position).Select(ToEmitParameter).ToArray(),
                PreserveParameterAttributes = representative.ParameterList.Any(static parameter => parameter.PreserveContractAttributes),
                Accessors = accessors,
                Documentation = FindDocumentation(documentation, DocumentationBindingKey(representative))
            };
        }
    }

    private static string EffectiveEmissionOwner(ProjectedType type, ProjectedMember member)
    {
        if (member.Invocation.OperationKind == ProjectionInvocationOperation.LocalForward)
        {
            var target = type.Members.FirstOrDefault(candidate =>
                candidate.Invocation.OperationKind != ProjectionInvocationOperation.LocalForward
                && candidate.CSharpName.Equals(member.Invocation.DispatchName, StringComparison.Ordinal)
                && candidate.ParameterList.Count == member.ParameterList.Count);
            if (target is not null)
                return string.IsNullOrWhiteSpace(target.EmissionOwnerLogicalId) ? type.LogicalId : target.EmissionOwnerLogicalId;
        }
        if (type.EventSink is not null
            && IsSinkOwner(member.EmissionOwnerLogicalId, type.EventSink.Name))
        {
            var matchingSinkMembers = type.Members.Where(candidate =>
                IsSinkOwner(candidate.EmissionOwnerLogicalId, type.EventSink.Name)
                && candidate.CSharpName.Equals(member.CSharpName, StringComparison.Ordinal)
                && ParametersMatch(candidate, member))
                .OrderByDescending(static candidate => candidate.ParameterList.Count(static parameter => parameter.SinkArgument is not null))
                .ThenByDescending(static candidate => candidate.Attributes.Count)
                .ThenBy(static candidate => candidate.IsContractDerived)
                .ThenBy(static candidate => candidate.EventSinkOnly)
                .ThenBy(static candidate => candidate.LogicalId, StringComparer.Ordinal)
                .ToArray();
            if (matchingSinkMembers.Length != 0 && !ReferenceEquals(matchingSinkMembers[0], member))
                return member.EmissionOwnerLogicalId;
            var alreadyOwnedByInterface = type.Members.Any(candidate =>
                !ReferenceEquals(candidate, member)
                && (string.IsNullOrWhiteSpace(candidate.EmissionOwnerLogicalId)
                    || candidate.EmissionOwnerLogicalId.Equals(type.LogicalId, StringComparison.Ordinal))
                && candidate.CSharpName.Equals(member.CSharpName, StringComparison.Ordinal)
                && ParametersMatch(candidate, member));
            if (!alreadyOwnedByInterface) return type.LogicalId;
        }
        return string.IsNullOrWhiteSpace(member.EmissionOwnerLogicalId) ? type.LogicalId : member.EmissionOwnerLogicalId;
    }

    private static bool IsSinkOwner(string ownerLogicalId, string sinkName)
        => ownerLogicalId.Equals(sinkName, StringComparison.Ordinal)
            || ownerLogicalId.EndsWith("." + sinkName, StringComparison.Ordinal);

    private static bool ParametersMatch(ProjectedMember left, ProjectedMember right)
        => left.ParameterList.Count == right.ParameterList.Count
            && left.ParameterList.OrderBy(static parameter => parameter.Position)
                .Zip(right.ParameterList.OrderBy(static parameter => parameter.Position),
                    static (left, right) => left.RefKind.Equals(right.RefKind, StringComparison.OrdinalIgnoreCase)
                        && SimpleTypeName(left.Type).Equals(SimpleTypeName(right.Type), StringComparison.Ordinal))
                .All(static matches => matches);

    private static ProjectedMember ApplyLocalForwardContract(ProjectedType type, ProjectedMember member)
    {
        if (member.Invocation.OperationKind != ProjectionInvocationOperation.LocalForward) return member;
        var candidates = type.ExactContractMembers.Where(candidate =>
            candidate.Name.Equals(member.CSharpName, StringComparison.Ordinal)
            && candidate.ParameterList.Count == member.ParameterList.Count).ToArray();
        var exact = candidates.FirstOrDefault(candidate =>
            candidate.ParameterList.OrderBy(static parameter => parameter.Position)
                .Zip(member.ParameterList.OrderBy(static parameter => parameter.Position),
                    static (left, right) => SimpleTypeName(left.Type).Equals(SimpleTypeName(right.Type), StringComparison.Ordinal))
                .All(static matches => matches))
            ?? (candidates.Length == 1 ? candidates[0] : null);
        return exact?.ReturnType is { Length: > 0 } returnType
            ? member with { ReturnType = returnType }
            : member;
    }

    private static string SimpleTypeName(string type)
    {
        var normalized = NormalizeManagedType(type);
        var generic = normalized.IndexOf('<');
        var prefix = generic < 0 ? normalized : normalized.Substring(0, generic);
        var separator = prefix.LastIndexOf('.');
        return (separator < 0 ? prefix : prefix.Substring(separator + 1)) + (generic < 0 ? "" : normalized.Substring(generic));
    }

    private static EmitMember ToEmitMember(ProjectedType owner, ProjectedMember member, DocumentationResult? documentation, bool containingConstants)
    {
        var modifiers = member.Modifiers.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var kind = containingConstants ? "constant" : NormalizeMemberKind(member.Kind);
        var isValue = kind is "enum-value" or "constant";
        var result = new EmitMember
        {
            Name = member.CSharpName,
            Kind = kind,
            Accessibility = member.Accessibility,
            Type = MemberType(member, kind),
            Static = modifiers.Contains("static"),
            Virtual = modifiers.Contains("virtual"),
            Override = modifiers.Contains("override"),
            Abstract = modifiers.Contains("abstract"),
            ReadOnly = modifiers.Contains("readonly"),
            Const = modifiers.Contains("const") || kind == "constant",
            New = modifiers.Contains("new"),
            Attributes = NormalizeAttributes(member.Attributes),
            Parameters = member.ParameterList.OrderBy(static parameter => parameter.Position).Select(ToEmitParameter).ToArray(),
            PreserveParameterAttributes = member.ParameterList.Any(static parameter => parameter.PreserveContractAttributes),
            DeclarationOrder = member.Line,
            UseCustomEventAccessors = member.UseCustomEventAccessors,
            EventBackingField = member.UseCustomEventAccessors
                ? member.EventBackingField ?? "_" + member.CSharpName
                : member.EventBackingField,
            EventValidationKey = member.EventValidationKey,
            EventValidationReleaseMode = member.EventInvalidReleaseArguments is null
                ? WrapperEventValidationReleaseMode.Suppress
                : WrapperEventValidationReleaseMode.Explicit,
            EventValidationReleaseArguments = member.EventInvalidReleaseArguments?.ToArray(),
            EventValidationInlineReturn = member.EventValidationInlineReturn,
            EventSinkOnly = member.EventSinkOnly,
            Value = isValue ? member.ConstantExpression ?? member.ConstantValue ?? ContractConstantExpression(member.Signature) ?? "0" : null,
            Documentation = FindDocumentation(documentation, DocumentationBindingKey(member))
        };
        if (member.RuntimeMemberKind == ProjectionRuntimeMemberKind.Clone)
            result.Body = $"return base.Clone() as {result.Type};";
        else if (member.RuntimeMemberKind == ProjectionRuntimeMemberKind.FromProxyService)
            result.Declaration = $"{member.Accessibility} {result.Type} {result.Name} {{ get; private set; }}";
        else if (member.RuntimeMemberKind is ProjectionRuntimeMemberKind.Dispose or ProjectionRuntimeMemberKind.DisposeWithEventBinding)
        {
            var argument = member.RuntimeMemberKind == ProjectionRuntimeMemberKind.DisposeWithEventBinding
                ? member.ParameterList.OrderBy(static parameter => parameter.Position).FirstOrDefault()?.CSharpName
                : null;
            result.Body = "if(this.Equals(GlobalHelperModules.GlobalModule.Instance))\n"
                + "    GlobalHelperModules.GlobalModule.Instance = null;\n"
                + "base.Dispose(" + (argument ?? string.Empty) + ");";
        }
        else if (kind == "method" && modifiers.Contains("static") && TryActiveInstanceBody(owner, result.Type, member.CSharpName, out var activeInstanceBody))
            result.Body = activeInstanceBody;
        else if (kind is "method")
            result.Invocation = MapInvocation(member, owner.Namespace);
        else if (kind is "property" or "indexer")
        {
            var accessorInvocations = member.AccessorInvocations.Count != 0
                ? member.AccessorInvocations
                : new[] { member.Invocation };
            result.Accessors = accessorInvocations
                .GroupBy(static invocation => EmitAccessorKind(invocation), StringComparer.Ordinal)
                .Select(static candidates => candidates.First())
                .OrderBy(static invocation => AccessorOrder(invocation))
                .Select(invocation => new EmitAccessor
                {
                    Kind = EmitAccessorKind(invocation),
                    Invocation = MapInvocation(member with { Invocation = invocation }, owner.Namespace)
                }).ToArray();
        }
        return result;
    }

    private static bool TryActiveInstanceBody(ProjectedType owner, string returnType, string memberName, out string body)
    {
        var product = JsonSerializer.Serialize(owner.Product);
        var className = JsonSerializer.Serialize(owner.CSharpName);
        if (memberName.Equals("GetActiveInstance", StringComparison.Ordinal))
        {
            body = $"return Running.ProxyService.GetActiveInstance<{returnType}>({product}, {className}, throwExceptionIfNotFound);";
            return true;
        }
        if (memberName.Equals("GetActiveInstances", StringComparison.Ordinal))
        {
            var open = returnType.IndexOf('<');
            var close = returnType.LastIndexOf('>');
            var itemType = open >= 0 && close > open ? returnType.Substring(open + 1, close - open - 1) : owner.CSharpName;
            body = $"return Running.ProxyService.GetActiveInstances<{itemType}>({product}, {className});";
            return true;
        }
        body = string.Empty;
        return false;
    }

    private static EmitMember ToEmitConstructor(ProjectedType type, ConstructorPlan constructor, DocumentationResult? documentation)
        => ToEmitConstructor(type.CSharpName, constructor, documentation);

    private static EmitMember ToEmitConstructor(string typeName, ConstructorPlan constructor, DocumentationResult? documentation)
    {
        var hidden = constructor.Kind is not "proxy-share" and not "proxy";
        return new EmitMember
        {
            Name = typeName,
            Kind = "constructor",
            Accessibility = string.IsNullOrWhiteSpace(constructor.Accessibility) ? "public" : constructor.Accessibility,
            Type = "void",
            Initializer = constructor.BaseCall,
            Parameters = constructor.Parameters.OrderBy(static parameter => parameter.Position).Select(ToEmitParameter).ToArray(),
            PreserveParameterAttributes = constructor.Parameters.Any(static parameter => parameter.PreserveContractAttributes),
            Attributes = constructor.IsContractOverlay || constructor.Attributes.Count != 0
                ? NormalizeAttributes(constructor.Attributes)
                : hidden
                    ? new[] { "Browsable(false)", "EditorBrowsable(EditorBrowsableState.Never)" }
                    : Array.Empty<string>(),
            Documentation = FindDocumentation(documentation, constructor.DocsBindingKey),
            Body = string.Empty
        };
    }

    private static WrapperInvocation? MapInvocation(ProjectedMember member, string ownerNamespace)
    {
        if (member.Invocation.OperationKind is ProjectionInvocationOperation.Field or ProjectionInvocationOperation.None or ProjectionInvocationOperation.Constructor) return null;
        var parameters = member.ParameterList.ToDictionary(static parameter => parameter.CSharpName, StringComparer.Ordinal);
        var arguments = member.Invocation.Arguments.Count != 0
            ? member.Invocation.Arguments.Select(argument =>
            {
                var expression = argument.Expression;
                var initializationExpression = argument.InitializationExpression;
                var parameterExpression = expression.StartsWith("(object)", StringComparison.Ordinal)
                    ? expression.Substring("(object)".Length).Trim()
                    : expression;
                if (string.IsNullOrWhiteSpace(initializationExpression)
                    && parameters.TryGetValue(parameterExpression, out var parameter)
                    && parameter.RefKind.Equals("out", StringComparison.OrdinalIgnoreCase))
                {
                    if (!parameterExpression.Equals(expression, StringComparison.Ordinal))
                        initializationExpression = "null";
                    else
                        initializationExpression = "default(" + ProjectedManagedType(parameter.TypeReference, parameter.Type) + ")";
                }
                return new WrapperInvocationArgument
                {
                    Expression = expression,
                    WriteBackExpression = argument.WriteBackExpression,
                    WriteBackType = argument.WriteBackType,
                    ByRef = argument.ByRef,
                    IsPropertyValue = argument.IsPropertyValue,
                    InitializationExpression = initializationExpression,
                    WriteBackConversion = argument.WriteBackConversion
                };
            }).ToList()
            : member.Invocation.ArgumentOrder.Select(name =>
            {
                if (name == "value")
                    return new WrapperInvocationArgument { Expression = name, IsPropertyValue = true };
                if (!parameters.TryGetValue(name, out var parameter))
                    return new WrapperInvocationArgument { Expression = name };
                var isOut = parameter.RefKind.Equals("out", StringComparison.OrdinalIgnoreCase);
                var writeBack = isOut || parameter.RefKind.Equals("ref", StringComparison.OrdinalIgnoreCase);
                return new WrapperInvocationArgument
                {
                    Expression = isOut ? "null" : parameter.CSharpName,
                    WriteBackExpression = writeBack ? parameter.CSharpName : null,
                    WriteBackType = ProjectedManagedType(parameter.TypeReference, parameter.Type)
                };
            }).ToList();
        var redirectTarget = RedirectTarget(member.Attributes);
        var kind = redirectTarget is not null
            ? WrapperInvocationKind.LocalForward
            : member.Invocation.CallKind switch
            {
                ProjectionInvocationCallKind.MethodGet => WrapperInvocationKind.Method,
                ProjectionInvocationCallKind.PropertyGet => WrapperInvocationKind.PropertyGet,
                ProjectionInvocationCallKind.PropertySet => WrapperInvocationKind.PropertySet,
                ProjectionInvocationCallKind.ValuePropertySet => WrapperInvocationKind.PropertySetValue,
                ProjectionInvocationCallKind.VariantPropertySet => WrapperInvocationKind.PropertySetVariant,
                ProjectionInvocationCallKind.EnumPropertySet => WrapperInvocationKind.PropertySetEnum,
                ProjectionInvocationCallKind.ReferencePropertySet => WrapperInvocationKind.PropertyPutRef,
                _ => member.Invocation.OperationKind switch
                {
                    ProjectionInvocationOperation.PropertyGet => WrapperInvocationKind.PropertyGet,
                    ProjectionInvocationOperation.PropertySetVariant => WrapperInvocationKind.PropertySetVariant,
                    ProjectionInvocationOperation.PropertySetEnum => WrapperInvocationKind.PropertySetEnum,
                    ProjectionInvocationOperation.PropertySet => WrapperInvocationKind.PropertySet,
                    ProjectionInvocationOperation.PropertyPutRef => WrapperInvocationKind.PropertyPutRef,
                    ProjectionInvocationOperation.LocalForward => WrapperInvocationKind.LocalForward,
                    ProjectionInvocationOperation.Event or ProjectionInvocationOperation.EventRaise => WrapperInvocationKind.EventRaise,
                    _ => WrapperInvocationKind.Method
                }
            };
        if (kind == WrapperInvocationKind.LocalForward)
        {
            for (var index = 0; index < arguments.Count && index < member.Invocation.ArgumentOrder.Count; index++)
            {
                if (!parameters.TryGetValue(member.Invocation.ArgumentOrder[index], out var parameter)) continue;
                var prefix = parameter.RefKind.Equals("out", StringComparison.OrdinalIgnoreCase)
                    ? "out "
                    : parameter.RefKind.Equals("ref", StringComparison.OrdinalIgnoreCase) ? "ref " : null;
                if (prefix is not null && !arguments[index].Expression.StartsWith(prefix, StringComparison.Ordinal))
                    arguments[index].Expression = prefix + arguments[index].Expression;
            }
        }
        if (kind is WrapperInvocationKind.PropertySet or WrapperInvocationKind.PropertySetValue or WrapperInvocationKind.PropertySetVariant or WrapperInvocationKind.PropertySetEnum or WrapperInvocationKind.PropertyPutRef
            && arguments.Count == 0)
            throw new InvalidOperationException("Projected property setter has no value argument: " + member.LogicalId
                + "; member-kind=" + member.Kind
                + "; argument-order=" + string.Join(",", member.Invocation.ArgumentOrder)
                + "; projected-arguments=" + member.Invocation.Arguments.Count.ToString(CultureInfo.InvariantCulture)
                + "; accessor-plans=" + member.AccessorInvocations.Count.ToString(CultureInfo.InvariantCulture));
        var exactReturn = kind == WrapperInvocationKind.LocalForward
            || member.Invocation.HasContractInvocation
            || member.Invocation.ReturnConversion is ProjectionReturnConversion.UntypedReference or ProjectionReturnConversion.Variant;
        var returnType = ProjectedManagedType(exactReturn ? null : member.Invocation.ResultTypeReference,
            member.ReturnType ?? member.Invocation.ResultType ?? "void");
        var returnKind = kind is WrapperInvocationKind.PropertySet or WrapperInvocationKind.PropertySetValue or WrapperInvocationKind.PropertySetVariant or WrapperInvocationKind.PropertySetEnum or WrapperInvocationKind.PropertyPutRef
            ? WrapperReturnKind.Void
            : MapReturnKind(member.Invocation.ReturnConversion, returnType);
        if (member.Invocation.InvokerCallStyle == ProjectionInvokerCallStyle.Direct)
            returnKind = returnType.Equals("void", StringComparison.Ordinal) ? WrapperReturnKind.Void : WrapperReturnKind.Raw;
        var needsInvoker = member.Invocation.Api == ProjectionInvocationApi.Invoker;
        return new WrapperInvocation
        {
            Api = needsInvoker ? WrapperInvocationApi.Invoker : WrapperInvocationApi.Factory,
            Kind = kind,
            ReturnKind = returnKind,
            Target = kind == WrapperInvocationKind.LocalForward ? null : member.Invocation.Target,
            DispatchName = redirectTarget ?? member.Invocation.DispatchName,
            ReturnType = returnKind == WrapperReturnKind.Void ? null : returnType,
            WrapperTypeExpression = returnKind == WrapperReturnKind.KnownReference
                ? WrapperType(member.Invocation.ResultTypeReference, returnType, ownerNamespace) + ".LateBindingApiWrapperType"
                : null,
            ScalarConversion = returnKind == WrapperReturnKind.Scalar
                ? !needsInvoker
                    ? FactoryScalarSuffix(member.Invocation.FactoryMethodSuffix ?? returnType)
                    : !InvokerSupportsScalar(returnType)
                        ? "(" + returnType + "){0}"
                        : null
                : null,
            ReleaseArguments = member.Invocation.ReleaseArguments,
            ArgumentPacking = member.Invocation.ArgumentPacking == ProjectionArgumentPacking.ObjectArray
                ? WrapperInvocationArgumentPacking.ObjectArray
                : WrapperInvocationArgumentPacking.Flat,
            ObjectArrayStyle = member.Invocation.ObjectArrayStyle == ProjectionObjectArrayStyle.Compact
                ? WrapperObjectArrayStyle.Compact
                : WrapperObjectArrayStyle.Spaced,
            RawCallCast = member.Invocation.RawCallCast == ProjectionRawCallCast.Object
                ? WrapperRawCallCast.Object
                : WrapperRawCallCast.None,
            KnownReferenceFactoryStyle = member.Invocation.KnownReferenceFactoryStyle == ProjectionKnownReferenceFactoryStyle.NonGeneric
                ? WrapperKnownReferenceFactoryStyle.NonGeneric
                : WrapperKnownReferenceFactoryStyle.Generic,
            ReturnValueStyle = member.Invocation.ReturnValueStyle == ProjectionReturnValueStyle.Local
                ? WrapperReturnValueStyle.Local
                : WrapperReturnValueStyle.Direct,
            ReturnCastStyle = member.Invocation.ReturnCastStyle switch
            {
                ProjectionReturnCastStyle.Explicit => WrapperReturnCastStyle.Explicit,
                ProjectionReturnCastStyle.As => WrapperReturnCastStyle.As,
                _ => WrapperReturnCastStyle.Default
            },
            ReturnLocalName = member.Invocation.LocalReturnName,
            ReturnLocalType = member.Invocation.LocalReturnType,
            InvokerCallStyle = member.Invocation.InvokerCallStyle == ProjectionInvokerCallStyle.Direct
                ? WrapperInvokerCallStyle.Direct
                : WrapperInvokerCallStyle.Standard,
            Arguments = arguments
        };
    }

    private static string WrapperType(ProjectedTypeReference? reference, string returnType, string ownerNamespace)
    {
        if (returnType.Contains('.', StringComparison.Ordinal) || returnType.Contains('<', StringComparison.Ordinal)) return returnType;
        var referencedType = ProjectedManagedType(reference, returnType);
        return SimpleTypeName(referencedType).Equals(SimpleTypeName(returnType), StringComparison.Ordinal)
            ? QualifyWrapperType(referencedType, ownerNamespace)
            : QualifyWrapperType(returnType, ownerNamespace);
    }

    private static string QualifyWrapperType(string type, string ownerNamespace)
        => type.Contains('.', StringComparison.Ordinal) || type.Contains('<', StringComparison.Ordinal) || string.IsNullOrWhiteSpace(ownerNamespace)
            ? type
            : "global::" + ownerNamespace + "." + type;

    private static string? RedirectTarget(IEnumerable<string> attributes)
    {
        foreach (var attribute in attributes)
        {
            if (!attribute.StartsWith("Redirect(", StringComparison.Ordinal)) continue;
            var firstQuote = attribute.IndexOf('"');
            var lastQuote = attribute.LastIndexOf('"');
            if (firstQuote >= 0 && lastQuote > firstQuote) return attribute.Substring(firstQuote + 1, lastQuote - firstQuote - 1);
        }
        return null;
    }

    private static WrapperReturnKind MapReturnKind(ProjectionReturnConversion conversion, string returnType)
        => conversion switch
        {
            ProjectionReturnConversion.None when returnType.Equals("void", StringComparison.OrdinalIgnoreCase) => WrapperReturnKind.Void,
            ProjectionReturnConversion.None => WrapperReturnKind.Raw,
            ProjectionReturnConversion.Scalar or ProjectionReturnConversion.String => WrapperReturnKind.Scalar,
            ProjectionReturnConversion.Enum => WrapperReturnKind.Enum,
            ProjectionReturnConversion.Value or ProjectionReturnConversion.Struct => WrapperReturnKind.Struct,
            ProjectionReturnConversion.Variant => WrapperReturnKind.Variant,
            ProjectionReturnConversion.KnownReference => WrapperReturnKind.KnownReference,
            ProjectionReturnConversion.Reference => WrapperReturnKind.Reference,
            ProjectionReturnConversion.UntypedReference => WrapperReturnKind.UntypedReference,
            ProjectionReturnConversion.BaseReference => WrapperReturnKind.BaseReference,
            ProjectionReturnConversion.EventArgument => WrapperReturnKind.EventArgument,
            ProjectionReturnConversion.Native or ProjectionReturnConversion.Array => WrapperReturnKind.Raw,
            _ => WrapperReturnKind.Raw
        };

    private static string FactoryScalarSuffix(string type)
    {
        var normalized = type.Replace("global::", string.Empty, StringComparison.Ordinal).Replace("System.", string.Empty, StringComparison.Ordinal);
        return normalized switch
        {
            "bool" or "Boolean" => "Bool",
            "byte" or "Byte" => "Byte",
            "short" or "Int16" => "Int16",
            "int" or "Int32" => "Int32",
            "long" or "Int64" => "Int64",
            "float" or "Single" => "Single",
            "double" or "Double" => "Double",
            "string" or "String" => "String",
            "DateTime" => "DateTime",
            _ => normalized
        };
    }

    private static bool FactorySupportsScalar(string type, WrapperInvocationKind kind)
    {
        var normalized = type.Replace("global::", string.Empty, StringComparison.Ordinal).Replace("System.", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
        if (kind == WrapperInvocationKind.Method && normalized is "boolean" or "bool") return false;
        if (kind == WrapperInvocationKind.Method)
            return normalized is "int16" or "short" or "int32" or "int" or "double" or "single" or "float" or "boolean" or "bool" or "datetime" or "string";
        return normalized is "boolean" or "bool" or "byte" or "int16" or "short" or "int32" or "int" or "int64" or "long" or "single" or "float" or "double" or "string" or "datetime";
    }

    private static bool InvokerSupportsScalar(string type)
    {
        var normalized = type.Replace("global::", string.Empty, StringComparison.Ordinal).Replace("System.", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
        return normalized is "boolean" or "bool" or "byte" or "int16" or "short" or "int32" or "int" or "int64" or "long" or "single" or "float" or "double" or "decimal" or "string" or "datetime";
    }

    private static EmitParameter ToEmitParameter(ProjectedParameter parameter)
    {
        var type = ProjectedManagedType(parameter.TypeReference, parameter.Type);
        var defaultValue = parameter.EmitDefaultValue && parameter.HasDefaultValue && parameter.RefKind.Equals("value", StringComparison.OrdinalIgnoreCase)
            ? parameter.IsContractOverlay
                ? parameter.DefaultValue
                : NormalizeDefaultValue(type, parameter.DefaultValue, parameter.TypeReference?.IsEnum == true)
            : null;
        var attributes = parameter.PreserveContractAttributes
            ? NormalizeAttributes(parameter.Attributes)
            : NormalizeAttributes(parameter.Attributes.Concat(ParameterAttributes(parameter.TypeReference)).Distinct(StringComparer.Ordinal));
        return new EmitParameter
        {
            Name = parameter.CSharpName,
            Type = type,
            DefaultValue = defaultValue,
            Ref = parameter.RefKind.Equals("ref", StringComparison.OrdinalIgnoreCase),
            Out = parameter.RefKind.Equals("out", StringComparison.OrdinalIgnoreCase),
            In = parameter.RefKind.Equals("in", StringComparison.OrdinalIgnoreCase),
            Attributes = attributes.ToArray(),
            EventConversion = MapEventConversion(parameter.SinkArgument)
        };
    }
    private static string[] NormalizeAttributes(IEnumerable<string> attributes)
        => attributes.Select(static attribute =>
            global::System.Text.RegularExpressions.Regex.Replace(
                attribute,
                @"(?<![A-Za-z0-9_:])System\.",
                "global::System.",
                global::System.Text.RegularExpressions.RegexOptions.CultureInvariant)).ToArray();

    private static WrapperEventArgumentConversion? MapEventConversion(EventSinkArgumentPlan? plan)
        => plan is null
            ? null
            : new WrapperEventArgumentConversion
            {
                Kind = plan.Conversion switch
                {
                    ProjectionEventConversion.Scalar => WrapperEventArgumentKind.Scalar,
                    ProjectionEventConversion.Enum => WrapperEventArgumentKind.Enum,
                    ProjectionEventConversion.KnownReference => WrapperEventArgumentKind.KnownReference,
                    ProjectionEventConversion.EventReference => WrapperEventArgumentKind.EventReference,
                    _ => WrapperEventArgumentKind.Raw
                },
                ManagedType = plan.ManagedType,
                WrapperTypeExpression = plan.WrapperTypeExpression,
                ConversionExpression = plan.ConversionExpression,
                WriteBackExpression = plan.WriteBackExpression,
                SourceArgument = plan.SourceArgument,
                LocalName = plan.LocalName
            };


    private static string NormalizeDefaultValue(string type, string? value, bool isEnum = false)
    {
        var normalized = value?.Trim() ?? string.Empty;
        if (type.Equals("string", StringComparison.OrdinalIgnoreCase) || type.Equals("System.String", StringComparison.OrdinalIgnoreCase) || type.Equals("global::System.String", StringComparison.OrdinalIgnoreCase))
        {
            if (normalized.Length >= 2 && normalized.StartsWith("\"", StringComparison.Ordinal) && normalized.EndsWith("\"", StringComparison.Ordinal)) return normalized;
            return JsonSerializer.Serialize(normalized);
        }
        if (normalized.Length == 0 || normalized.StartsWith("null ", StringComparison.OrdinalIgnoreCase) || normalized.Equals("Nothing", StringComparison.OrdinalIgnoreCase)) return "null";
        if (isEnum && (normalized[0] == '-' || char.IsDigit(normalized[0]))) return $"({type})({normalized})";
        return normalized;
    }

    private static EmitDocumentation? FindDocumentation(DocumentationResult? result, string bindingKey)
    {
        if (result is null || !result.Documents.TryGetValue(bindingKey, out var bound)) return null;
        var document = bound.Documentation;
        var examples = document.Elements.Where(static element => element.Name.Equals("example", StringComparison.OrdinalIgnoreCase)).Select(ElementText).ToArray();
        var exceptions = document.Elements.Where(static element => element.Name.Equals("exception", StringComparison.OrdinalIgnoreCase))
            .Select(static element => element.Attributes.TryGetValue("cref", out var value) ? value : null)
            .Where(static value => !string.IsNullOrWhiteSpace(value)).Cast<string>().ToArray();
        return new EmitDocumentation
        {
            RawXml = document.RawXml,
            Summary = document.Summary,
            Remarks = document.Remarks,
            Returns = document.Returns,
            Value = document.Value,
            Parameters = new Dictionary<string, string>(document.Parameters, StringComparer.Ordinal),
            Examples = examples,
            Exceptions = exceptions
        };
    }

    private static string ElementText(XmlDocumentationElement element)
    {
        try { return XElement.Parse(element.Xml).Value.Trim(); }
        catch { return element.Xml; }
    }

    private static string DocumentationBindingKey(ProjectedMember member)
        => member.OverloadOrdinal == 0 ? member.DocsBindingKey : member.DocsBindingKey + ":overload:" + member.OverloadOrdinal.ToString(CultureInfo.InvariantCulture);

    private static string AccessorProjectionKey(ProjectedMember member)
        => member.AccessorGroupId + "|" + member.Kind + "|" + string.Join(",", member.ParameterList.OrderBy(static parameter => parameter.Position).Select(static parameter => parameter.Type + " " + parameter.CSharpName));

    private static string ProjectionMemberSignatureKey(ProjectedMember member)
        => member.Kind + "|" + member.CSharpName + "|"
            + (member.Kind is "property" or "indexer" ? EmitAccessorKind(member.Invocation) : "") + "|"
            + string.Join(",", member.ParameterList.OrderBy(static parameter => parameter.Position)
                .Select(static parameter => parameter.RefKind + ":" + SimpleTypeName(parameter.Type)));

    private static int AccessorOrder(InvocationPlan invocation)
        => EmitAccessorKind(invocation) == "get" ? 0 : 1;

    private static string EmitAccessorKind(InvocationPlan invocation)
        => invocation.CallKind switch
        {
            ProjectionInvocationCallKind.MethodGet or ProjectionInvocationCallKind.PropertyGet => "get",
            ProjectionInvocationCallKind.PropertySet
                or ProjectionInvocationCallKind.ValuePropertySet
                or ProjectionInvocationCallKind.VariantPropertySet
                or ProjectionInvocationCallKind.EnumPropertySet
                or ProjectionInvocationCallKind.ReferencePropertySet => "set",
            _ => invocation.OperationKind == ProjectionInvocationOperation.PropertyGet ? "get" : "set"
        };

    private static string NormalizeTypeKind(string kind, ProjectionEntityKind entityKind)
    {
        var normalized = kind.Trim().ToLowerInvariant();
        if (entityKind == ProjectionEntityKind.Module) return "module";
        if (entityKind == ProjectionEntityKind.Constants) return "constants";
        return normalized switch
        {
            "coclass" => "class",
            "record" => "struct",
            "typedef" or "alias" => "struct",
            "dispatch" or "dispatchinterface" => "class",
            _ => normalized
        };
    }

    private static string NormalizeMemberKind(string kind)
        => kind.Trim().ToLowerInvariant() switch
        {
            "enum" or "enumvalue" => "enum-value",
            "const" => "constant",
            var value => value
        };

    private static string? ContractConstantExpression(string signature)
    {
        var equals = signature.IndexOf('=', StringComparison.Ordinal);
        if (equals < 0) return null;
        var expression = signature.Substring(equals + 1).Trim().TrimEnd(';', ',').Trim();
        return expression.Length == 0 ? null : expression;
    }

    private static string MemberType(ProjectedMember member, string kind)
    {
        if (kind == "enum-value") return NormalizeManagedType(member.ValueType ?? "int");
        if (kind == "constant") return NormalizeManagedType(member.ValueType ?? InferConstantType(member.ConstantExpression ?? member.ConstantValue));
        return ProjectedManagedType(member.Invocation.OperationKind == ProjectionInvocationOperation.LocalForward || member.Invocation.HasContractInvocation ? null : member.Invocation.ResultTypeReference, member.ReturnType ?? member.ValueType ?? "object");
    }

    private static string ProjectedManagedType(ProjectedTypeReference? reference, string fallback)
    {
        var type = string.IsNullOrWhiteSpace(reference?.QualifiedName) ? fallback : reference.QualifiedName;
        if (reference?.IsArray == true && !type.EndsWith("[]", StringComparison.Ordinal)) type += "[]";
        return NormalizeManagedType(type);
    }

    private static IReadOnlyList<string> ParameterAttributes(ProjectedTypeReference? reference)
    {
        if (string.IsNullOrWhiteSpace(reference?.MarshalAs)) return Array.Empty<string>();
        var value = reference.MarshalAs.StartsWith("System.Runtime.InteropServices.", StringComparison.Ordinal)
            ? "global::" + reference.MarshalAs
            : reference.MarshalAs.StartsWith("global::System.Runtime.InteropServices.", StringComparison.Ordinal)
                ? reference.MarshalAs
                : "global::System.Runtime.InteropServices." + reference.MarshalAs;
        return new[] { $"global::System.Runtime.InteropServices.MarshalAs({value})" };
    }

    private static string NormalizeManagedType(string type)
    {
        var normalized = type.Trim();
        return normalized.Equals("COMVariant", StringComparison.Ordinal) ? "object" : normalized;
    }


    private static string InferConstantType(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "int";
        if (value.StartsWith('"')) return "string";
        if (value is "true" or "false") return "bool";
        if (value.Contains('.', StringComparison.Ordinal)) return "double";
        return "int";
    }

    private static void ValidateGenerationMode(CommandRequest request)
    {
        ValidateReadMode(request);
        if (request.Apply) throw new ArgumentException("--apply is only valid for bootstrap-ownership.");
        if (request.IsolatedOutput && string.IsNullOrWhiteSpace(request.SourcePath)) throw new ArgumentException("--source is required by --isolated-output.");
    }

    private static void ValidateReadMode(CommandRequest request)
    {
        if (request.FromData && request.Locked) throw new ArgumentException("--from-data/--exploratory is exploratory only and cannot be used with --locked.");
        if (request.Locked && string.IsNullOrWhiteSpace(request.DataPath)) throw new ArgumentException("--data is required in locked mode.");
        if (string.IsNullOrWhiteSpace(request.DataPath)) throw new ArgumentException("--data is required.");
        if (string.IsNullOrWhiteSpace(request.PolicyPath)) throw new ArgumentException("--policy is required.");
    }

    private static GraphInput LoadGraph(string? root)
    {
        var location = Required(root, "--data");
        var path = File.Exists(location) ? location : Path.Combine(location, "graph.json");
        if (!File.Exists(path)) throw new ArgumentException($"Canonical Data v2 graph is required: {path}");
        using var hashStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, FileOptions.SequentialScan);
        var fileHash = Convert.ToHexString(SHA256.HashData(hashStream)).ToLowerInvariant();
        return new GraphInput(CanonicalJson.Read(path), fileHash);
    }

    private static IReadOnlyList<ContractInput> LoadContracts(string? path, IReadOnlyList<string> products)
    {
        if (string.IsNullOrWhiteSpace(path)) return Array.Empty<ContractInput>();
        var fullPath = Path.GetFullPath(path);
        string root;
        string[] files;
        if (File.Exists(fullPath))
        {
            root = Path.GetDirectoryName(fullPath)!;
            files = new[] { fullPath };
        }
        else if (Directory.Exists(fullPath))
        {
            root = fullPath;
            files = Directory.EnumerateFiles(fullPath, "*.wrapper-contract.json", SearchOption.AllDirectories)
                .Where(file => products.Contains(Path.GetFileName(file)[..^".wrapper-contract.json".Length], StringComparer.OrdinalIgnoreCase))
                .Order(StringComparer.Ordinal).ToArray();
        }
        else throw new FileNotFoundException($"Wrapper Contract path is missing: {fullPath}", fullPath);
        if (files.Length == 0) throw new InvalidDataException($"No *.wrapper-contract.json files were found beneath '{fullPath}'.");
        return files.Select(file =>
        {
            var bytes = File.ReadAllBytes(file);
            return new ContractInput(Path.GetRelativePath(root, file).Replace('\\', '/'), bytes, HashBytes(bytes), ContractApi(bytes), InvalidDocumentationKeys(bytes));
        }).ToArray();
    }

    private static string ContractApi(byte[] bytes)
    {
        using var document = JsonDocument.Parse(bytes);
        if (document.RootElement.TryGetProperty("Source", out var source)
            && source.TryGetProperty("Api", out var api)
            && api.GetString() is { Length: > 0 } value) return value;
        throw new InvalidDataException("Wrapper Contract Source.Api is required.");
    }

    private static IReadOnlySet<string> InvalidDocumentationKeys(byte[] bytes)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        using var document = JsonDocument.Parse(bytes);
        if (!document.RootElement.TryGetProperty("Types", out var types) || types.ValueKind != JsonValueKind.Array) return result;
        foreach (var type in types.EnumerateArray())
        {
            if (!type.TryGetProperty("LogicalId", out var typeIdElement) || typeIdElement.GetString() is not { Length: > 0 } typeId) continue;
            if (HasInvalidDocumentation(type)) result.Add("T\u0000" + typeId);
            if (!type.TryGetProperty("Members", out var members) || members.ValueKind != JsonValueKind.Array) continue;
            foreach (var member in members.EnumerateArray())
            {
                if (!HasInvalidDocumentation(member) || !member.TryGetProperty("Signature", out var signatureElement) || signatureElement.GetString() is not { } signature) continue;
                result.Add("M\u0000" + typeId + "\u0000" + signature);
                result.Add("S\u0000" + signature);
            }
        }
        return result;

        static bool HasInvalidDocumentation(JsonElement declaration)
            => declaration.TryGetProperty("Documentation", out var documentation)
               && documentation.ValueKind == JsonValueKind.Object
               && documentation.TryGetProperty("ParseStatus", out var status)
               && status.GetString()?.Equals("invalid", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static WrapperContract SyntheticContract()
        => new()
        {
            SchemaVersion = "1.0",
            ContractKind = "NetOffice.WrapperContract",
            Generator = new ContractGenerator { Name = "NetOffice.CodeGen exploratory Data v2", Version = "1" },
            Source = new ContractSource { Root = "Data v2", Api = "__data__", SourceRoots = new[] { "Data v2" } },
            Types = Array.Empty<ContractType>(),
            Unknowns = Array.Empty<ContractUnknown>(),
            Ambiguities = Array.Empty<ContractAmbiguity>()
        };


    private static IReadOnlyList<string> SelectProducts(DataGraph graph, string? specification)
    {
        var available = graph.Projects.Select(static project => project.Name).Distinct(StringComparer.OrdinalIgnoreCase).Order(StringComparer.Ordinal).ToArray();
        var defaults = graph.Projects.Where(static project => !project.Ignore).Select(static project => project.Name).Distinct(StringComparer.OrdinalIgnoreCase).Order(StringComparer.Ordinal).ToArray();
        if (string.IsNullOrWhiteSpace(specification) || specification.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Any(static value => value.Equals("all", StringComparison.OrdinalIgnoreCase))) return defaults;
        var requested = specification.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (requested.Length == 0) throw new ArgumentException("--projects must name at least one product or 'all'.");
        var unknown = requested.Where(value => !available.Contains(value, StringComparer.OrdinalIgnoreCase)).Order(StringComparer.Ordinal).ToArray();
        if (unknown.Length != 0) throw new ArgumentException($"Unknown product(s): {string.Join(", ", unknown)}. Available products: {string.Join(", ", available)}.");
        return available.Where(product => requested.Contains(product, StringComparer.OrdinalIgnoreCase)).ToArray();
    }

    private static DataGraph FilterGraphForProduct(DataGraph graph, string product)
    {
        var projectIds = graph.Projects.Where(project => project.Name.Equals(product, StringComparison.OrdinalIgnoreCase))
            .Select(static project => project.LogicalId).ToHashSet(StringComparer.Ordinal);
        var typesById = graph.Types.ToDictionary(static type => type.LogicalId, StringComparer.Ordinal);
        var membersByType = graph.Members.ToLookup(static member => member.TypeId, StringComparer.Ordinal);
        var includedTypeIds = graph.Types.Where(type => projectIds.Contains(type.ProjectId)).Select(static type => type.LogicalId).ToHashSet(StringComparer.Ordinal);
        var pending = new Queue<string>(includedTypeIds);
        while (pending.Count != 0)
        {
            var typeId = pending.Dequeue();
            if (!typesById.TryGetValue(typeId, out var type)) continue;
            foreach (var reference in type.BaseTypeIds.Concat(type.DefaultInterfaceIds).Concat(type.EventInterfaceIds).Concat(type.ReferenceObservations.Select(static observation => observation.TargetTypeId)))
                AddType(reference);
            if (!projectIds.Contains(type.ProjectId)) continue;
            foreach (var member in membersByType[typeId])
            {
                AddType(member.ReturnTypeReference?.TargetTypeId);
                foreach (var parameter in member.Parameters) AddType(parameter.TypeReference?.TargetTypeId);
            }
        }

        var members = graph.Members.Where(member => typesById.TryGetValue(member.TypeId, out var owner) && projectIds.Contains(owner.ProjectId)).ToArray();
        var memberIds = members.Select(static member => member.LogicalId).ToHashSet(StringComparer.Ordinal);
        var values = graph.Values.Where(value => typesById.TryGetValue(value.TypeId, out var owner) && projectIds.Contains(owner.ProjectId)).ToArray();
        var valueIds = values.Select(static value => value.LogicalId).ToHashSet(StringComparer.Ordinal);
        var accessorGroups = graph.AccessorGroups.Where(group => memberIds.Overlaps(group.MemberIds)).ToArray();
        var accessorGroupIds = accessorGroups.Select(static group => group.LogicalId).ToHashSet(StringComparer.Ordinal);
        var allTargetIds = includedTypeIds.Concat(memberIds).Concat(valueIds).ToHashSet(StringComparer.Ordinal);
        var filtered = graph with
        {
            Types = graph.Types.Where(type => includedTypeIds.Contains(type.LogicalId)).ToArray(),
            Members = members,
            Values = values,
            AccessorGroups = accessorGroups,
            SupportObservations = graph.SupportObservations.Where(observation => allTargetIds.Contains(observation.TargetId)).ToArray(),
            Aliases = graph.Aliases.Where(alias => allTargetIds.Contains(alias.TargetId)).ToArray(),
            Unifications = graph.Unifications.Where(unification => allTargetIds.Contains(unification.CanonicalId) && unification.EquivalentIds.All(allTargetIds.Contains)).ToArray(),
            InvocationEvidence = graph.InvocationEvidence.Where(evidence => memberIds.Contains(evidence.MemberId)).ToArray(),
            AccessorEvidence = graph.AccessorEvidence.Where(evidence => accessorGroupIds.Contains(evidence.AccessorGroupId) && evidence.MemberIds.All(memberIds.Contains)).ToArray(),
            AbsentFacts = graph.AbsentFacts.Where(record => record.TargetId is null || allTargetIds.Contains(record.TargetId)).ToArray(),
            Unknowns = graph.Unknowns.Where(record => allTargetIds.Contains(record.LogicalId)).ToArray(),
            StaleRecords = graph.StaleRecords.Where(record => allTargetIds.Contains(record.TargetId)).ToArray(),
            Ambiguities = graph.Ambiguities.Where(record => allTargetIds.Contains(record.LogicalId)).ToArray(),
            Digest = string.Empty
        };
        return filtered with { Digest = CanonicalJson.ComputeDigest(filtered) };

        void AddType(string? id)
        {
            if (!string.IsNullOrWhiteSpace(id) && includedTypeIds.Add(id)) pending.Enqueue(id);
        }
    }

    private static IReadOnlyList<CompanionFile> CollectCompanions(string sourceRoot, string outputRoot, IReadOnlyList<string> products, string? contractPath)
    {
        if (PathsOverlap(sourceRoot, outputRoot)) throw new ArgumentException("--isolated-output must not overlap --source; source tree mutation is forbidden.");
        var result = new Dictionary<string, CompanionFile>(StringComparer.Ordinal);
        var classifiedSources = LoadClassifiedWrapperSources(contractPath);
        AddRepositoryFile(Path.Combine(Path.GetDirectoryName(sourceRoot)!, "LICENSE.txt"), "LICENSE.txt", "build-metadata");
        AddRepositoryFile(Path.Combine(Path.GetDirectoryName(sourceRoot)!, "icon.png"), "icon.png", "build-metadata");
        foreach (var name in new[] { "NetOffice.props", "NetOffice.snk", "trustedsigning.json" })
            AddSourceFile(Path.Combine(sourceRoot, name), name, "build-metadata");
        AddTree(Path.Combine(sourceRoot, "NetOffice"), "NetOffice", static _ => "runtime-companion", excludeGeneratedCategories: false);
        foreach (var product in products)
        {
            var projectRoot = Path.Combine(sourceRoot, product);
            if (!Directory.Exists(projectRoot)) continue;
            AddTree(projectRoot, product, relative => ClassifyProductCompanion(product, relative, classifiedSources), excludeGeneratedCategories: true);
        }
        return result.Values.OrderBy(static file => file.RelativePath, StringComparer.Ordinal).ToArray();

        void AddTree(string root, string destination, Func<string, string?> classifier, bool excludeGeneratedCategories)
        {
            if (!Directory.Exists(root)) return;
            foreach (var file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
            {
                var relative = Path.GetRelativePath(root, file).Replace('\\', '/');
                if (IsTransientPath(relative)) continue;
                if (excludeGeneratedCategories && IsGeneratedWrapperSource(Path.GetFileName(root), relative, classifiedSources)) continue;
                var classification = classifier(relative);
                if (classification is null) continue;
                Add(file, $"Source/{destination}/{relative}", classification);
            }
        }
        void AddSourceFile(string path, string destination, string classification)
        {
            if (File.Exists(path)) Add(path, "Source/" + destination, classification);
        }
        void AddRepositoryFile(string path, string destination, string classification)
        {
            if (File.Exists(path)) Add(path, destination, classification);
        }
        void Add(string source, string destination, string classification)
        {
            var bytes = File.ReadAllBytes(source);
            result[destination] = new CompanionFile(destination, bytes, classification, source);
        }
    }

    private static IReadOnlyDictionary<string, Dictionary<string, string>> LoadClassifiedWrapperSources(string? contractPath)
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(contractPath)) return result;
        var fullPath = Path.GetFullPath(contractPath);
        var root = File.Exists(fullPath) ? Path.GetDirectoryName(fullPath)! : fullPath;
        if (!Directory.Exists(root)) return result;
        foreach (var file in Directory.EnumerateFiles(root, "*.classification.json", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
        {
            using var document = JsonDocument.Parse(File.ReadAllBytes(file));
            if (!document.RootElement.TryGetProperty("Api", out var apiElement) || apiElement.GetString() is not { Length: > 0 } api) continue;
            if (!result.TryGetValue(api, out var paths)) result[api] = paths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (document.RootElement.TryGetProperty("Files", out var files) && files.ValueKind == JsonValueKind.Array)
            {
                foreach (var entry in files.EnumerateArray())
                {
                    if (!entry.TryGetProperty("Path", out var path) || path.GetString() is not { Length: > 0 } value) continue;
                    var generate = entry.TryGetProperty("BuildAction", out var action) && action.GetString()?.Equals("generate", StringComparison.OrdinalIgnoreCase) == true;
                    var ownership = entry.TryGetProperty("Ownership", out var owner) ? owner.GetString() : null;
                    paths[value.Replace('\\', '/')] = generate ? "generated-wrapper" : ownership?.Equals("runtime", StringComparison.OrdinalIgnoreCase) == true ? "runtime-companion" : "manual-companion";
                }
            }
            else if (document.RootElement.TryGetProperty("Entries", out var entries) && entries.ValueKind == JsonValueKind.Array)
            {
                foreach (var entry in entries.EnumerateArray())
                    if (entry.TryGetProperty("Source", out var source) && source.GetString() is { Length: > 0 } value) paths[value.Replace('\\', '/')] = "generated-wrapper";
            }
        }
        return result;
    }

    private static bool IsGeneratedWrapperSource(string product, string relative, IReadOnlyDictionary<string, Dictionary<string, string>> classified)
    {
        var normalized = relative.Replace('\\', '/');
        if (classified.TryGetValue(product, out var paths) && paths.TryGetValue(normalized, out var classification)) return classification == "generated-wrapper";
        var first = normalized.Split('/')[0];
        return GeneratedCategoryNames.Contains(first, StringComparer.OrdinalIgnoreCase) && normalized.EndsWith(".cs", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ClassifyProductCompanion(string product, string relative, IReadOnlyDictionary<string, Dictionary<string, string>> classified)
    {
        var normalized = relative.Replace('\\', '/');
        if (classified.TryGetValue(product, out var paths) && paths.TryGetValue(normalized, out var classification))
            return classification == "generated-wrapper" ? null : classification;
        return IsGeneratedWrapperSource(product, normalized, classified) ? null : normalized.EndsWith(".cs", StringComparison.OrdinalIgnoreCase) ? "manual-companion" : "build-metadata";
    }

    private static bool IsTransientPath(string relative)
        => relative.Split('/').Any(segment => segment.Equals("bin", StringComparison.OrdinalIgnoreCase) || segment.Equals("obj", StringComparison.OrdinalIgnoreCase) || segment.Equals(".vs", StringComparison.OrdinalIgnoreCase));

    private static IReadOnlyList<string> FindCompanionChanges(string outputRoot, IReadOnlyList<CompanionFile> companions)
        => companions.Where(file => !File.Exists(Path.Combine(outputRoot, file.RelativePath.Replace('/', Path.DirectorySeparatorChar)))
                || !File.ReadAllBytes(Path.Combine(outputRoot, file.RelativePath.Replace('/', Path.DirectorySeparatorChar))).AsSpan().SequenceEqual(file.Bytes))
            .Select(static file => file.RelativePath).Order(StringComparer.Ordinal).ToArray();

    private static void WriteCompanions(string outputRoot, IReadOnlyList<CompanionFile> companions, CancellationToken cancellationToken)
    {
        foreach (var companion in companions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var destination = Path.Combine(outputRoot, companion.RelativePath.Replace('/', Path.DirectorySeparatorChar));
            if (File.Exists(destination) && File.ReadAllBytes(destination).AsSpan().SequenceEqual(companion.Bytes)) continue;
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            var temporary = destination + ".codegen-" + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllBytes(temporary, companion.Bytes);
                File.Move(temporary, destination, true);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }
    }

    private static void EnsureUniqueDesiredPaths(IReadOnlyList<EmittedFile> emitted, IReadOnlyList<CompanionFile> companions)
    {
        var duplicates = emitted.Select(static file => file.RelativePath).Concat(companions.Select(static file => file.RelativePath))
            .GroupBy(static path => path, StringComparer.Ordinal).Where(static group => group.Count() > 1).Select(static group => group.Key).Order(StringComparer.Ordinal).ToArray();
        if (duplicates.Length != 0) throw new InvalidDataException($"Emitted and companion paths collide: {string.Join(", ", duplicates)}.");
    }

    private static CommandReportDetails CreateDetails(ProjectionInputs inputs, DocumentationResult? documentation, IReadOnlyList<EmittedFile> emitted, IReadOnlyList<CompanionFile> companions, string? treeHash)
    {
        var hashes = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["graph"] = inputs.GraphDigest,
            ["graphFile"] = inputs.GraphFileHash,
            ["policy"] = inputs.PolicyDigest,
            ["contracts"] = inputs.ContractAggregateHash
        };
        if (documentation is not null) hashes["documentation"] = documentation.Digest;
        return new CommandReportDetails
        {
            InputHashes = hashes,
            ContractFiles = inputs.ContractHashes,
            EntityCounts = inputs.Counts,
            Products = inputs.Products,
            EmittedPaths = emitted.Select(static file => file.RelativePath).Order(StringComparer.Ordinal).ToArray(),
            CopiedPaths = companions.Select(static file => new CopiedPathReport(file.RelativePath, file.Classification, file.SourcePath)).OrderBy(static file => file.Path, StringComparer.Ordinal).ToArray(),
            OutputTreeHash = treeHash,
            CacheUsed = false
        };
    }

    private static CommandResult Diff(CommandRequest request)
    {
        var expected = ExistingDirectory(request.ExpectedPath ?? request.SourcePath, "--expected/--source");
        var actual = ExistingDirectory(request.ActualPath ?? request.OutputPath, "--actual/--output");
        var expectedFiles = FileHashes(expected);
        var actualFiles = FileHashes(actual);
        var differences = expectedFiles.Keys.Union(actualFiles.Keys)
            .Where(path => !expectedFiles.TryGetValue(path, out var left) || !actualFiles.TryGetValue(path, out var right) || left != right)
            .Order(StringComparer.Ordinal).ToArray();
        return differences.Length == 0
            ? new CommandResult(0, "No differences.", differences)
            : new CommandResult(2, $"Differences: {string.Join(", ", differences)}", differences);
    }

    private static Dictionary<string, string> FileHashes(string root)
        => Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Where(path => !Path.GetRelativePath(root, path).Replace('\\', '/').StartsWith(".codegen/", StringComparison.Ordinal))
            .ToDictionary(path => Path.GetRelativePath(root, path).Replace('\\', '/'), path => HashBytes(File.ReadAllBytes(path)), StringComparer.Ordinal);

    private static void WriteReport(string? path, CommandRequest request, CommandResult result, long elapsedMilliseconds, long peakManagedBytes)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        var reportPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(reportPath)!);
        var details = result.Details;
        var report = new
        {
            schemaVersion = "codegen-report-v3",
            command = request.Command,
            mode = request.FromData ? "exploratory" : request.Locked ? "locked" : "unlocked",
            isolatedOutput = request.IsolatedOutput,
            exitCode = result.ExitCode,
            message = result.Message,
            inputHashes = details?.InputHashes ?? new Dictionary<string, string>(),
            contractFiles = details?.ContractFiles ?? new Dictionary<string, string>(),
            entityCounts = details?.EntityCounts ?? new Dictionary<string, long>(),
            products = details?.Products ?? Array.Empty<string>(),
            changedPaths = result.ChangedPaths.Order(StringComparer.Ordinal),
            emittedPaths = details?.EmittedPaths ?? Array.Empty<string>(),
            copiedPaths = details?.CopiedPaths ?? Array.Empty<CopiedPathReport>(),
            outputTreeHash = details?.OutputTreeHash,
            cacheUsed = details?.CacheUsed ?? false,
            traces = details?.Traces ?? Array.Empty<ProjectionTrace>(),
            elapsedMilliseconds = details?.ElapsedMilliseconds ?? elapsedMilliseconds,
            peakManagedBytes = details?.PeakManagedBytes ?? peakManagedBytes
        };
        File.WriteAllText(reportPath, JsonSerializer.Serialize(report, ReportJsonOptions) + Environment.NewLine, new UTF8Encoding(false));
    }

    private static CommandResult Fail(string message) => new(1, message, Array.Empty<string>());

    private static string Required(string? path, string option)
        => string.IsNullOrWhiteSpace(path) ? throw new ArgumentException($"{option} is required.") : Path.GetFullPath(path);

    private static string ExistingDirectory(string? path, string option)
    {
        var result = Required(path, option);
        if (!Directory.Exists(result)) throw new ArgumentException($"{option} must name an existing directory.");
        return result;
    }

    private static string ExistingFile(string? path, string option)
    {
        var result = Required(path, option);
        if (!File.Exists(result)) throw new ArgumentException($"{option} must name an existing file.");
        return result;
    }

    private static OwnershipManifest? ReadManifest(string root)
    {
        var path = Path.Combine(root, ".codegen", "ownership.json");
        return File.Exists(path) ? OwnershipManifest.Parse(File.ReadAllBytes(path)) : null;
    }

    private static string SafePathPart(string value)
    {
        var invalid = Path.GetInvalidFileNameChars().ToHashSet();
        var result = new string(value.Select(character => invalid.Contains(character) || character is '/' or '\\' ? '_' : character).ToArray()).Trim();
        if (result.Length == 0 || result is "." or "..") throw new InvalidDataException($"Unsafe product path '{value}'.");
        return result;
    }

    private static string HashBytes(ReadOnlySpan<byte> bytes)
        => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static bool PathsOverlap(string first, string second)
    {
        var left = Path.GetFullPath(first).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var right = Path.GetFullPath(second).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return left.Equals(right, StringComparison.OrdinalIgnoreCase)
            || left.StartsWith(right + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            || right.StartsWith(left + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }

    private static readonly JsonSerializerOptions ReportJsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private sealed record GraphInput(DataGraph Graph, string FileHash);
    private sealed record ContractInput(string RelativePath, byte[] Bytes, string Sha256, string Api, IReadOnlySet<string> InvalidDocumentationKeys);
    private sealed record CompanionFile(string RelativePath, byte[] Bytes, string Classification, string SourcePath);
    private sealed record ProjectionInputs(
        string GraphDigest,
        string GraphFileHash,
        string PolicyDigest,
        string ContractAggregateHash,
        IReadOnlyDictionary<string, string> ContractHashes,
        IReadOnlyList<string> Products,
        IReadOnlyList<ProjectedFile> Files,
        IReadOnlyList<ProjectionTrace> Traces,
        IReadOnlyDictionary<string, long> Counts);

    private static void CollectStageGarbage()
        => GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: false);

    private sealed class ManagedMemorySampler : IDisposable
    {
        private readonly Timer _timer;
        private long _peakBytes;

        public ManagedMemorySampler()
        {
            _peakBytes = GC.GetTotalMemory(false);
            _timer = new Timer(_ => Sample(), null, TimeSpan.Zero, TimeSpan.FromMilliseconds(20));
        }

        public long PeakBytes
        {
            get
            {
                Sample();
                return Interlocked.Read(ref _peakBytes);
            }
        }

        public void Dispose()
        {
            _timer.Dispose();
            Sample();
        }

        private void Sample()
        {
            var value = GC.GetTotalMemory(false);
            var current = Interlocked.Read(ref _peakBytes);
            while (value > current)
            {
                var observed = Interlocked.CompareExchange(ref _peakBytes, value, current);
                if (observed == current) break;
                current = observed;
            }
        }
    }
}
