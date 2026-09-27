using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace NetOffice.CodeGen.Storage
{
    public static class ContentHasher
    {
        public static string Sha256(byte[] bytes)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            using (var sha = SHA256.Create()) return string.Concat(sha.ComputeHash(bytes).Select(x => x.ToString("x2")));
        }
    }

    public interface IContentCache
    {
        bool TryGet(string sha256, out byte[] bytes);
        void Put(string sha256, byte[] bytes);
    }

    public sealed class ContentCache : IContentCache
    {
        private readonly ConcurrentDictionary<string, byte[]> memory = new ConcurrentDictionary<string, byte[]>(StringComparer.Ordinal);
        private readonly string root;
        public ContentCache(string root = null) { this.root = string.IsNullOrWhiteSpace(root) ? null : Path.GetFullPath(root); }

        public bool TryGet(string sha256, out byte[] bytes)
        {
            ValidateKey(sha256);
            if (memory.TryGetValue(sha256, out bytes)) { bytes = (byte[])bytes.Clone(); return true; }
            if (root == null) { bytes = null; return false; }
            var path = Path.Combine(root, sha256.Substring(0, 2), sha256 + ".bin");
            if (!File.Exists(path)) { bytes = null; return false; }
            var loaded = File.ReadAllBytes(path);
            if (!string.Equals(ContentHasher.Sha256(loaded), sha256, StringComparison.Ordinal)) { bytes = null; return false; }
            memory[sha256] = loaded;
            bytes = (byte[])loaded.Clone();
            return true;
        }

        public void Put(string sha256, byte[] bytes)
        {
            ValidateKey(sha256);
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (!string.Equals(ContentHasher.Sha256(bytes), sha256, StringComparison.Ordinal)) throw new ArgumentException("Content does not match its SHA-256 key.");
            var copy = (byte[])bytes.Clone();
            memory[sha256] = copy;
            if (root == null) return;
            var path = Path.Combine(root, sha256.Substring(0, 2), sha256 + ".bin");
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            if (File.Exists(path)) return;
            var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllBytes(temporary, copy);
            try { File.Move(temporary, path); } catch (IOException) { if (File.Exists(path)) File.Delete(temporary); else throw; }
        }

        private static void ValidateKey(string key)
        {
            if (key == null || key.Length != 64 || key.Any(x => !((x >= '0' && x <= '9') || (x >= 'a' && x <= 'f')))) throw new ArgumentException("A lowercase SHA-256 key is required.", nameof(key));
        }
    }
}
