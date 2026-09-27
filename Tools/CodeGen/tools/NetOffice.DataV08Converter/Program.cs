using System.Text.Json;
using System.Text.Json.Serialization;
using NetOffice.CodeGen.Data;

return await ConverterProgram.RunAsync(args);

internal static class ConverterProgram
{
    private static readonly JsonSerializerOptions ReportOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true
    };

    public static async Task<int> RunAsync(string[] args)
    {
        var options = ParseOptions(args);
        if (options.Error is not null)
            return await WriteFailureAsync(options.Report, "usage", options.Error);
        var inputOption = options.Input!;
        var outputOption = options.Output!;
        var reportOption = options.Report!;
        var schemaOption = options.Schema!;
        if (!File.Exists(schemaOption))
            return await WriteFailureAsync(reportOption, "schema-missing", $"Schema file was not found: {schemaOption}");

        var input = ResolveInput(inputOption);
        if (input.Error is not null)
            return await WriteFailureAsync(reportOption, "input-unavailable", input.Error);

        try
        {
            var bytes = await File.ReadAllBytesAsync(input.Path!);
            var parsed = CanonicalInputConverter.Parse(System.Text.Encoding.UTF8.GetString(bytes));
            var source = parsed.Source ?? new DataSource();
            if (string.IsNullOrWhiteSpace(source.Sha256))
                source = source with { Sha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant() };
            if (string.IsNullOrWhiteSpace(source.Path))
                source = source with { Path = "canonical-input" };
            parsed = parsed with { Source = source };

            var conversion = CanonicalInputConverter.Convert(parsed);
            var validation = DataGraphValidator.Validate(conversion.Graph);
            var issues = conversion.Issues.Concat(validation.Issues).Select(static issue => issue.ToString()).ToArray();
            var graphPath = Path.Combine(outputOption, "graph.json");
            if (issues.Length == 0)
            {
                Directory.CreateDirectory(outputOption);
                CanonicalJson.Write(graphPath, conversion.Graph);
            }

            var report = new
            {
                schemaVersion = DataSchema.Version,
                serialization = DataSchema.Serialization,
                input = input.Path,
                output = issues.Length == 0 ? graphPath : null,
                status = issues.Length == 0 ? "ok" : "invalid",
                graphDigest = issues.Length == 0 ? conversion.Graph.Digest : null,
                issues
            };
            await WriteReportAsync(reportOption, report);
            return issues.Length == 0 ? 0 : 1;
        }
        catch (JsonException exception)
        {
            return await WriteFailureAsync(reportOption, "input-invalid", exception.Message);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            return await WriteFailureAsync(reportOption, "conversion-failed", exception.Message);
        }
    }

    private static (string? Input, string? Output, string? Report, string? Schema, string? Error) ParseOptions(string[] args)
    {
        string? input = null;
        string? output = null;
        string? report = null;
        string? schema = null;
        for (var index = 0; index < args.Length; index++)
        {
            var value = args[index];
            if (value is "--input" or "--output" or "--report" or "--schema")
            {
                if (++index >= args.Length)
                    return (null, null, report, schema ?? "schema.json", $"Missing value for {value}.");
                switch (value)
                {
                    case "--input": input = args[index]; break;
                    case "--output": output = args[index]; break;
                    case "--report": report = args[index]; break;
                    case "--schema": schema = args[index]; break;
                }
            }
            else
                return (null, null, report, schema ?? "schema.json", $"Unknown argument {value}.");
        }
        if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output) || string.IsNullOrWhiteSpace(report))
            return (null, null, report, schema ?? "schema.json", "--input, --output, and --report are required.");
        return (input, output, report, schema ?? "schema.json", null);
    }

    private static (string? Path, string? Error) ResolveInput(string input)
    {
        if (File.Exists(input))
            return (input, null);
        if (!Directory.Exists(input))
            return (null, $"Input path does not exist: {input}");
        var canonical = Directory.EnumerateFiles(input, "*.json", SearchOption.TopDirectoryOnly)
            .Where(static path => string.Equals(Path.GetFileName(path), "canonical-input.json", StringComparison.OrdinalIgnoreCase))
            .OrderBy(static path => path, StringComparer.Ordinal)
            .FirstOrDefault();
        if (canonical is not null)
            return (canonical, null);
        return (null, $"No canonical-input.json found in {input}. XML/v0.8 input is intentionally unsupported; supply the v1 canonical input format.");
    }

    private static async Task<int> WriteFailureAsync(string? path, string status, string error)
    {
        if (path is not null)
            await WriteReportAsync(path, new { schemaVersion = DataSchema.Version, status, issues = new[] { error } });
        return 2;
    }

    private static async Task WriteReportAsync(string path, object report)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        await File.WriteAllTextAsync(fullPath, JsonSerializer.Serialize(report, ReportOptions) + "\n");
    }
}
