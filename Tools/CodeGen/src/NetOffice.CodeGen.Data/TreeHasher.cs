using System.Security.Cryptography;
using System.Text;

namespace NetOffice.CodeGen.Data;

public static class TreeHasher
{
    public static TreeHashResult ComputeDirectory(string directory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);
        var root = Path.GetFullPath(directory);
        if (!Directory.Exists(root))
            throw new DirectoryNotFoundException(root);

        var files = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Select(path => (Path: path, Relative: NormalizeRelativePath(root, path)))
            .OrderBy(item => item.Relative, StringComparer.Ordinal)
            .ToArray();
        var entries = files.Select(item => (item.Relative, Digest: ComputeFile(item.Path))).ToArray();
        return new TreeHashResult(DataSchema.DigestAlgorithm, ComputeEntriesDigest(entries), entries.Select(static item => item.Relative).ToArray());
    }

    public static string ComputeFiles(IEnumerable<(string RelativePath, ReadOnlyMemory<byte> Content)> files)
    {
        ArgumentNullException.ThrowIfNull(files);
        var entries = files
            .Select(item => (RelativePath: NormalizeRelativePath(item.RelativePath), Digest: ComputeBytes(item.Content.Span)))
            .OrderBy(item => item.RelativePath, StringComparer.Ordinal)
            .ToArray();
        return ComputeEntriesDigest(entries);
    }

    public static string ComputeFile(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    }

    private static string ComputeEntriesDigest(IEnumerable<(string RelativePath, string Digest)> entries)
    {
        var builder = new StringBuilder();
        foreach (var entry in entries)
        {
            builder.Append(entry.RelativePath).Append('\0').Append(entry.Digest).Append('\n');
        }
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString()))).ToLowerInvariant();
    }

    private static string ComputeBytes(ReadOnlySpan<byte> bytes)
        => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string NormalizeRelativePath(string root, string path)
        => NormalizeRelativePath(Path.GetRelativePath(root, path));

    private static string NormalizeRelativePath(string path)
        => path.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
}
