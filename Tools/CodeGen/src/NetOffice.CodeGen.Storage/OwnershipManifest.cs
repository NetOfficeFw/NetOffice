using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Storage
{
    public sealed class ManifestFile
    {
        [JsonPropertyName("path")] public string Path { get; set; }
        [JsonPropertyName("sha256")] public string Sha256 { get; set; }
        [JsonPropertyName("length")] public long Length { get; set; }
    }

    public sealed class OwnershipManifest
    {
        public const int CurrentSchemaVersion = 1;
        [JsonPropertyName("schemaVersion")] public int SchemaVersion { get; set; } = CurrentSchemaVersion;
        [JsonPropertyName("manifestKind")] public string ManifestKind { get; set; } = "netoffice-codegen-ownership";
        [JsonPropertyName("generatorVersion")] public string GeneratorVersion { get; set; } = "1";
        [JsonPropertyName("runHash")] public string RunHash { get; set; }
        [JsonPropertyName("files")] public IList<ManifestFile> Files { get; set; } = new List<ManifestFile>();

        public void Validate()
        {
            if (SchemaVersion != CurrentSchemaVersion) throw new InvalidDataException("Unsupported ownership manifest schema version: " + SchemaVersion);
            if (!string.Equals(ManifestKind, "netoffice-codegen-ownership", StringComparison.Ordinal)) throw new InvalidDataException("Invalid ownership manifest kind.");
            if (string.IsNullOrWhiteSpace(GeneratorVersion)) throw new InvalidDataException("Manifest generatorVersion is required.");
            if (Files == null) throw new InvalidDataException("Manifest files are required.");
            var paths = new HashSet<string>(StringComparer.Ordinal);
            foreach (var file in Files)
            {
                var path = PathPolicy.NormalizeRelative(file.Path);
                if (!paths.Add(path)) throw new InvalidDataException("Duplicate manifest path: " + path);
                if (!IsHash(file.Sha256)) throw new InvalidDataException("Invalid SHA-256 for " + path + ".");
                if (file.Length < 0) throw new InvalidDataException("Invalid file length for " + path + ".");
            }
            var expectedRunHash = ComputeRunHash(Files);
            if (!string.IsNullOrWhiteSpace(RunHash) && !string.Equals(RunHash, expectedRunHash, StringComparison.Ordinal)) throw new InvalidDataException("Ownership manifest runHash does not match its files.");
            RunHash = expectedRunHash;
        }

        public byte[] ToUtf8()
        {
            Validate();
            var options = new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = null, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.Default };
            var ordered = new OwnershipManifest { SchemaVersion = SchemaVersion, ManifestKind = ManifestKind, GeneratorVersion = GeneratorVersion, RunHash = RunHash, Files = Files.OrderBy(x => x.Path, StringComparer.Ordinal).ToList() };
            return new UTF8Encoding(false).GetBytes(JsonSerializer.Serialize(ordered, options) + "\n");
        }

        public static OwnershipManifest Parse(byte[] bytes)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            var manifest = JsonSerializer.Deserialize<OwnershipManifest>(bytes);
            if (manifest == null) throw new InvalidDataException("Empty ownership manifest.");
            manifest.Validate();
            return manifest;
        }

        public static OwnershipManifest Create(string generatorVersion, IEnumerable<ManifestFile> files)
        {
            var manifest = new OwnershipManifest { GeneratorVersion = string.IsNullOrWhiteSpace(generatorVersion) ? "1" : generatorVersion.Trim(), Files = (files ?? Enumerable.Empty<ManifestFile>()).Select(x => new ManifestFile { Path = PathPolicy.NormalizeRelative(x.Path), Sha256 = x.Sha256, Length = x.Length }).OrderBy(x => x.Path, StringComparer.Ordinal).ToList() };
            manifest.RunHash = ComputeRunHash(manifest.Files);
            manifest.Validate();
            return manifest;
        }

        internal static string ComputeRunHash(IEnumerable<ManifestFile> files)
        {
            var canonical = string.Join("\n", (files ?? Enumerable.Empty<ManifestFile>()).OrderBy(x => x.Path, StringComparer.Ordinal).Select(x => x.Path + "\0" + x.Sha256 + "\0" + x.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            return ContentHasher.Sha256(new UTF8Encoding(false).GetBytes(canonical));
        }

        private static bool IsHash(string hash) { return hash != null && hash.Length == 64 && hash.All(x => (x >= '0' && x <= '9') || (x >= 'a' && x <= 'f')); }
    }

    public static class PathPolicy
    {
        public static string NormalizeRelative(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new InvalidDataException("A relative path is required.");
            var normalized = path.Replace('\\', '/').TrimStart('/');
            if (normalized.Length == 0 || normalized.Split('/').Any(x => x == ".." || x.Length == 0)) throw new InvalidDataException("Path escapes the managed root: " + path);
            return normalized;
        }

        public static bool IsAllowlistedGeneratedPath(string relativePath)
        {
            var normalized = NormalizeRelative(relativePath);
            return normalized.Split('/').Any(x => string.Equals(x, "Generated", StringComparison.Ordinal));
        }
    }
}
