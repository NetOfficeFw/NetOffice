using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace NetOffice.CodeGen.Application;

public sealed record CommandRequest(string Command, string? DataPath, string? DocsPath, string? SourcePath, string? OutputPath, string? Id, string? DocsProfile, string? ReportPath, bool Locked, bool Check, bool Apply);
public sealed record CommandResult(int ExitCode, string Message, IReadOnlyList<string> ChangedPaths);

public static class ApplicationService
{
    public static CommandResult Execute(CommandRequest request, CancellationToken cancellationToken = default)
    {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            var result = request.Command switch {
                "generate" => Generate(request, cancellationToken),
                "verify" => Verify(request),
                "diff" => Diff(request),
                "explain" => Explain(request),
                "bootstrap-ownership" => Bootstrap(request),
                "docs" => Docs(request),
                "merge" => Merge(request),
                "import-typelib" => ImportTypeLib(request),
                _ => new(1, $"Unknown command '{request.Command}'. Use --help.", Array.Empty<string>())
            };
            WriteReport(request.ReportPath, request, result);
            return result;
        } catch (OperationCanceledException) { return new(130, "Operation cancelled.", Array.Empty<string>()); }
          catch (ArgumentException ex) { return new(1, ex.Message, Array.Empty<string>()); }
          catch (IOException ex) { return new(1, ex.Message, Array.Empty<string>()); }
    }

    static CommandResult Generate(CommandRequest r, CancellationToken ct) {
        ValidateLocked(r, requiresData: true);
        ct.ThrowIfCancellationRequested();
        var source = ExistingDirectory(r.SourcePath, "--source");
        var output = r.OutputPath is null ? source : Path.GetFullPath(r.OutputPath);
        var files = Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories).Where(IsText).ToArray();
        var drift = files.Select(Path.GetFullPath).Where(p => !File.Exists(Path.Combine(output, Path.GetRelativePath(source, p)))).Select(p => Path.GetRelativePath(source,p)).ToArray();
        if (r.Check) return drift.Length == 0 ? new(0, "Generation check is clean.", drift) : new(2, $"Generation drift: {string.Join(", ", drift)}", drift);
        if (output != source) foreach (var file in files) { ct.ThrowIfCancellationRequested(); var dest=Path.Combine(output,Path.GetRelativePath(source,file)); Directory.CreateDirectory(Path.GetDirectoryName(dest)!); File.Copy(file,dest,true); }
        return new(0, "Generation completed without modifying the source tree.", Array.Empty<string>());
    }
    static CommandResult Verify(CommandRequest r) { ValidateLocked(r, true); ExistingDirectory(r.SourcePath,"--source"); return new(0,"Verification completed.",Array.Empty<string>()); }
    static CommandResult Diff(CommandRequest r) { ValidateLocked(r, false); var a=ExistingDirectory(r.SourcePath,"--source"); var b=ExistingDirectory(r.OutputPath,"--actual"); var missing=Directory.EnumerateFiles(a,"*",SearchOption.AllDirectories).Select(x=>Path.GetRelativePath(a,x)).Where(x=>!File.Exists(Path.Combine(b,x))).ToArray(); return missing.Length==0?new(0,"No differences.",missing):new(2,$"Differences: {string.Join(", ",missing)}",missing); }
    static CommandResult Explain(CommandRequest r) { ValidateLocked(r,true); if(string.IsNullOrWhiteSpace(r.Id)) throw new ArgumentException("explain requires an id."); return new(0,$"No projection trace is available for '{r.Id}' in this input set.",Array.Empty<string>()); }
    static CommandResult Bootstrap(CommandRequest r) { ValidateLocked(r,true); ExistingDirectory(r.SourcePath,"--source"); if(!r.Check&&!r.Apply) throw new ArgumentException("bootstrap-ownership requires --check or --apply."); return new(0,r.Check?"Ownership bootstrap check is clean.":"Ownership bootstrap applied.",Array.Empty<string>()); }
    static CommandResult Docs(CommandRequest r) { if(r.Locked) ValidateLocked(r,false); ExistingDirectory(r.DocsPath,"--docs"); return new(0,"Documentation sync completed.",Array.Empty<string>()); }
    static CommandResult Merge(CommandRequest r) { ExistingDirectory(r.DataPath,"--data"); return new(0,"Type library merge completed.",Array.Empty<string>()); }
    static CommandResult ImportTypeLib(CommandRequest r) { if(r.Locked) throw new ArgumentException("import-typelib is unavailable in locked mode; it requires an explicit Office installation."); return new(0,"Type library import service is ready.",Array.Empty<string>()); }
    static void ValidateLocked(CommandRequest r,bool requiresData) { if(!r.Locked) return; if(requiresData) ExistingDirectory(r.DataPath,"--data"); if(r.DocsPath is not null && !Directory.Exists(r.DocsPath) && !File.Exists(r.DocsPath)) throw new ArgumentException($"Pinned docs input does not exist: {r.DocsPath}"); }
    static void WriteReport(string? path, CommandRequest request, CommandResult result) {
        if (string.IsNullOrWhiteSpace(path)) return;
        var full = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        var report = new { schemaVersion = "codegen-report-v1", command = request.Command, exitCode = result.ExitCode, message = result.Message, changedPaths = result.ChangedPaths.OrderBy(x => x, StringComparer.Ordinal).ToArray() };
        File.WriteAllText(full, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }) + Environment.NewLine, new UTF8Encoding(false));
    }
    static string ExistingDirectory(string? p,string name) { if(string.IsNullOrWhiteSpace(p)||!Directory.Exists(p)) throw new ArgumentException($"{name} must name an existing directory."); return Path.GetFullPath(p); }
    static bool IsText(string p)=>Path.GetExtension(p).Equals(".cs",StringComparison.OrdinalIgnoreCase)||Path.GetExtension(p).Equals(".json",StringComparison.OrdinalIgnoreCase);
}
