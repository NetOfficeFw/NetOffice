namespace NetOffice.CodeGen.Docs;

/// <summary>
/// A completed session produces the same canonical result and digest as one full baseline sync.
/// Loads one baseline Wrapper Contract batch at a time and retains only bound results.
/// </summary>
public sealed class BaselineDocumentationSession
{
    private readonly DocsOptions _options;
    private readonly string _configuredContractPath;
    private readonly List<MappingRecord> _mappings = new();
    private readonly Dictionary<string, BoundDocumentation> _documents = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _documentHashes = new(StringComparer.Ordinal);
    private readonly HashSet<string> _logicalIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _bindingKeys = new(StringComparer.Ordinal);
    private readonly bool _retainDocuments;
    private bool _completed;

    public BaselineDocumentationSession(DocsOptions options, bool retainDocuments = true)
    {
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        if (options.Profile != DocsProfile.Baseline)
            throw new ArgumentException("A baseline documentation session requires the baseline profile.", nameof(options));
        _options = options;
        _configuredContractPath = Path.GetFullPath(options.BaselineContractPath!);
        _retainDocuments = retainDocuments;
    }

    /// <summary>Binds one batch from the configured Wrapper Contract path and retains its records.</summary>
    public DocumentationResult BindBatch(IEnumerable<DocumentationTarget> targets)
        => BindBatch(targets, _configuredContractPath);

    /// <summary>Binds one batch from a contract file or subdirectory beneath the configured contract path.</summary>
    public DocumentationResult BindBatch(IEnumerable<DocumentationTarget> targets, string baselineContractPath)
    {
        ArgumentNullException.ThrowIfNull(targets);
        if (_completed) throw new InvalidOperationException("The baseline documentation session is complete.");
        var orderedTargets = DocumentationSync.OrderAndValidateTargets(targets);

        foreach (var target in orderedTargets)
        {
            if (_logicalIds.Contains(target.LogicalId))
                throw new InvalidDataException($"Duplicate documentation logical ID '{target.LogicalId}' across baseline batches.");
            if (_bindingKeys.Contains(target.BindingKey))
                throw new InvalidDataException($"Duplicate documentation binding key '{target.BindingKey}' across baseline batches.");
        }
        var batchContractPath = ValidateBatchContractPath(baselineContractPath);
        var corpus = BaselineContractCorpus.Load(batchContractPath);
        var result = DocumentationSync.BindBaselineBatch(orderedTargets, corpus);
        if (result.Mappings.Any(static mapping => mapping.Kind == MappingKind.Ambiguous))
            throw new DocumentationMappingException(result);

        foreach (var target in orderedTargets)
        {
            _logicalIds.Add(target.LogicalId);
            _bindingKeys.Add(target.BindingKey);
        }
        _mappings.AddRange(result.Mappings);
        foreach (var document in result.Documents)
        {
            if (_retainDocuments) _documents.Add(document.Key, document.Value);
            else _documentHashes.Add(document.Key, DocumentationSync.Hash(document.Value.Documentation.RawXml));
        }
        return result;
    }

    private string ValidateBatchContractPath(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var fullPath = Path.GetFullPath(path);
        if (File.Exists(_configuredContractPath))
        {
            if (!string.Equals(fullPath, _configuredContractPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("A batch contract must match the configured Wrapper Contract file.");
            return fullPath;
        }

        var relativePath = Path.GetRelativePath(_configuredContractPath, fullPath);
        if (Path.IsPathRooted(relativePath)
            || relativePath.Equals("..", StringComparison.Ordinal)
            || relativePath.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
            || relativePath.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal))
            throw new InvalidOperationException("A batch contract must be beneath the configured Wrapper Contract directory.");
        return fullPath;
    }

    /// <summary>Finalizes all batches into one globally canonical result. No further batches are accepted.</summary>
    public DocumentationResult Complete()
    {
        if (_completed) throw new InvalidOperationException("The baseline documentation session is already complete.");
        _completed = true;
        var result = _retainDocuments
            ? DocumentationSync.CreateResult(_mappings, _documents, null, DocsProfile.Baseline, canonicalize: true)
            : DocumentationSync.CreateSummaryResult(_mappings, _documentHashes, DocsProfile.Baseline);
        if (!string.IsNullOrWhiteSpace(_options.ReportDirectory))
            DocumentationSync.WriteReports(_options.ReportDirectory, result);
        if (result.Mappings.Any(static mapping => mapping.Kind == MappingKind.Ambiguous))
            throw new DocumentationMappingException(result);
        return result;
    }
}
