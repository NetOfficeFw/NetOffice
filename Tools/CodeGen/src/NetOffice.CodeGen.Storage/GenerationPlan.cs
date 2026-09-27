using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using NetOffice.CodeGen.Emit;
using System.Threading.Tasks;

namespace NetOffice.CodeGen.Storage
{
    public enum FileOperationKind { Write, Delete }

    public sealed class PlannedOperation
    {
        public FileOperationKind Kind { get; internal set; }
        public string RelativePath { get; internal set; }
        public byte[] Content { get; internal set; }
        public string ContentSha256 { get; internal set; }
        public string ExpectedCurrentSha256 { get; internal set; }
        public bool IsManifest { get; internal set; }
    }

    public sealed class GenerationPlan
    {
        private readonly string outputRoot;
        private readonly Action<int, PlannedOperation> beforeOperation;
        private readonly string manifestPath;
        public IReadOnlyList<PlannedOperation> Operations { get; }
        public OwnershipManifest Manifest { get; }
        public bool CheckOnly { get; }
        public bool HasChanges { get { return Operations.Count != 0; } }

        internal GenerationPlan(string outputRoot, string manifestPath, OwnershipManifest manifest, IReadOnlyList<PlannedOperation> operations, bool checkOnly, Action<int, PlannedOperation> beforeOperation)
        {
            this.outputRoot = outputRoot; this.manifestPath = manifestPath; Manifest = manifest; Operations = operations; CheckOnly = checkOnly; this.beforeOperation = beforeOperation;
        }

        /// <summary>Applies every operation or restores all touched files if any replacement fails.</summary>
        public void Apply()
        {
            if (CheckOnly) throw new InvalidOperationException("A check-only plan cannot write files.");
            if (!HasChanges) return;
            Directory.CreateDirectory(outputRoot);
            var before = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            foreach (var operation in Operations)
            {
                var full = FullPath(operation.RelativePath);
                before[full] = File.Exists(full) ? File.ReadAllBytes(full) : null;
                var actual = File.Exists(full) ? ContentHasher.Sha256(before[full]) : null;
                if (!string.Equals(actual, operation.ExpectedCurrentSha256, StringComparison.Ordinal)) throw new HashMismatchException(operation.RelativePath, operation.ExpectedCurrentSha256, actual);
            }
            var stage = Path.Combine(outputRoot, ".codegen-stage-" + Guid.NewGuid().ToString("N"));
            var changed = new List<string>();
            try
            {
                foreach (var operation in Operations.Where(x => x.Kind == FileOperationKind.Write))
                {
                    var staged = Path.Combine(stage, operation.RelativePath.Replace('/', Path.DirectorySeparatorChar));
                    Directory.CreateDirectory(Path.GetDirectoryName(staged));
                    File.WriteAllBytes(staged, operation.Content);
                }
                var operationIndex = 0;
                foreach (var operation in Operations)
                {
                    beforeOperation?.Invoke(operationIndex++, operation);
                    var full = FullPath(operation.RelativePath);
                    if (operation.Kind == FileOperationKind.Delete)
                    {
                        if (!PathPolicy.IsAllowlistedGeneratedPath(operation.RelativePath)) throw new InvalidOperationException("Deletion is restricted to Generated roots: " + operation.RelativePath);
                        if (File.Exists(full)) { File.Delete(full); changed.Add(full); }
                    }
                    else
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(full));
                        var staged = Path.Combine(stage, operation.RelativePath.Replace('/', Path.DirectorySeparatorChar));
                        File.Move(staged, full, true);
                        changed.Add(full);
                    }
                }
            }
            catch
            {
                for (var i = changed.Count - 1; i >= 0; i--)
                {
                    var path = changed[i];
                    if (before[path] == null) { if (File.Exists(path)) File.Delete(path); }
                    else File.WriteAllBytes(path, before[path]);
                }
                throw;
            }
            finally { if (Directory.Exists(stage)) Directory.Delete(stage, true); }
        }

        private string FullPath(string relativePath)
        {
            var normalized = PathPolicy.NormalizeRelative(relativePath);
            var full = Path.GetFullPath(Path.Combine(outputRoot, normalized.Replace('/', Path.DirectorySeparatorChar)));
            var root = outputRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Path escapes output root: " + relativePath);
            return full;
        }
    }

    public sealed class PlanOptions
    {
        public string ManifestRelativePath { get; set; } = ".codegen/ownership.json";
        public string GeneratorVersion { get; set; } = "1";
        public bool CheckOnly { get; set; }
        public int MaxDegreeOfParallelism { get; set; } = 1;
        public Action<int, PlannedOperation> BeforeOperation { get; set; }
    }

    public sealed class HashMismatchException : InvalidOperationException
    {
        public string RelativePath { get; }
        public string Expected { get; }
        public string Actual { get; }
        public HashMismatchException(string path, string expected, string actual) : base("Managed file changed outside the generator: " + path + " (expected " + (expected ?? "missing") + ", actual " + (actual ?? "missing") + ").") { RelativePath = path; Expected = expected; Actual = actual; }
    }

    public static class GenerationPlanBuilder
    {
        public static GenerationPlan Build(string outputRoot, IEnumerable<EmittedFile> desiredFiles, OwnershipManifest previousManifest = null, PlanOptions options = null)
        {
            if (string.IsNullOrWhiteSpace(outputRoot)) throw new ArgumentNullException(nameof(outputRoot));
            options = options ?? new PlanOptions();
            outputRoot = Path.GetFullPath(outputRoot);
            var desired = new List<EmittedFile>(desiredFiles ?? Enumerable.Empty<EmittedFile>());
            var byPath = new Dictionary<string, EmittedFile>(StringComparer.Ordinal);
            foreach (var file in desired)
            {
                var path = PathPolicy.NormalizeRelative(file.RelativePath);
                if (!byPath.TryAdd(path, file)) throw new InvalidDataException("Duplicate desired path: " + path);
                if (!string.Equals(ContentHasher.Sha256(file.Bytes), file.Sha256, StringComparison.Ordinal)) throw new InvalidDataException("Emitted content hash mismatch: " + path);
            }
            previousManifest?.Validate();
            var old = (previousManifest?.Files ?? new List<ManifestFile>()).Select(x => new ManifestFile { Path = PathPolicy.NormalizeRelative(x.Path), Sha256 = x.Sha256, Length = x.Length }).ToDictionary(x => x.Path, StringComparer.Ordinal);
            var manifestFiles = byPath.OrderBy(x => x.Key, StringComparer.Ordinal).Select(x => new ManifestFile { Path = x.Key, Sha256 = x.Value.Sha256, Length = x.Value.Bytes.LongLength }).ToList();
            var manifest = OwnershipManifest.Create(options.GeneratorVersion, manifestFiles);
            var parallel = Math.Max(1, options.MaxDegreeOfParallelism);
            var operations = new ConcurrentBag<PlannedOperation>();
            try
            {
                Parallel.ForEach(byPath, new ParallelOptions { MaxDegreeOfParallelism = parallel }, pair =>
                {
                    var path = pair.Key; var file = pair.Value; var full = FullPath(outputRoot, path);
                    var current = File.Exists(full) ? File.ReadAllBytes(full) : null;
                    var currentHash = current == null ? null : ContentHasher.Sha256(current);
                    ManifestFile oldFile;
                    if (old.TryGetValue(path, out oldFile) && !string.Equals(currentHash, oldFile.Sha256, StringComparison.Ordinal)) throw new HashMismatchException(path, oldFile.Sha256, currentHash);
                    if (!string.Equals(currentHash, file.Sha256, StringComparison.Ordinal)) operations.Add(new PlannedOperation { Kind = FileOperationKind.Write, RelativePath = path, Content = file.Bytes, ContentSha256 = file.Sha256, ExpectedCurrentSha256 = currentHash });
                });
            }
            catch (AggregateException error)
            {
                var first = error.Flatten().InnerExceptions.FirstOrDefault();
                if (first != null) throw first;
                throw;
            }
            foreach (var oldFile in old.Values.Where(x => !byPath.ContainsKey(x.Path)).OrderBy(x => x.Path, StringComparer.Ordinal))
            {
                if (!PathPolicy.IsAllowlistedGeneratedPath(oldFile.Path)) throw new InvalidOperationException("Deletion is restricted to Generated roots: " + oldFile.Path);
                var full = FullPath(outputRoot, oldFile.Path);
                var current = File.Exists(full) ? File.ReadAllBytes(full) : null;
                var currentHash = current == null ? null : ContentHasher.Sha256(current);
                if (!string.Equals(currentHash, oldFile.Sha256, StringComparison.Ordinal)) throw new HashMismatchException(oldFile.Path, oldFile.Sha256, currentHash);
                if (current != null) operations.Add(new PlannedOperation { Kind = FileOperationKind.Delete, RelativePath = oldFile.Path, ExpectedCurrentSha256 = currentHash });
            }
            var manifestPath = PathPolicy.NormalizeRelative(options.ManifestRelativePath);
            if (byPath.ContainsKey(manifestPath)) throw new InvalidDataException("Manifest path collides with a generated source path: " + manifestPath);
            var manifestBytes = manifest.ToUtf8();
            var manifestFull = FullPath(outputRoot, manifestPath);
            var manifestCurrent = File.Exists(manifestFull) ? File.ReadAllBytes(manifestFull) : null;
            var manifestCurrentHash = manifestCurrent == null ? null : ContentHasher.Sha256(manifestCurrent);
            if (!string.Equals(manifestCurrentHash, ContentHasher.Sha256(manifestBytes), StringComparison.Ordinal)) operations.Add(new PlannedOperation { Kind = FileOperationKind.Write, RelativePath = manifestPath, Content = manifestBytes, ContentSha256 = ContentHasher.Sha256(manifestBytes), ExpectedCurrentSha256 = manifestCurrentHash, IsManifest = true });
            var ordered = operations.OrderBy(x => x.IsManifest ? 1 : 0).ThenBy(x => x.Kind == FileOperationKind.Delete ? 1 : 0).ThenBy(x => x.RelativePath, StringComparer.Ordinal).ToList();
            return new GenerationPlan(outputRoot, manifestPath, manifest, ordered, options.CheckOnly, options.BeforeOperation);
        }

        private static string FullPath(string root, string relative)
        {
            var full = Path.GetFullPath(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
            var prefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Path escapes output root: " + relative);
            return full;
        }
    }
}
