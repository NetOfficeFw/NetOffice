using System.Security.Cryptography;
using System.Text;

namespace NetOffice.CodeGen.Data;

public static class LogicalIds
{
    public static string Project(string sourceKey) => Create("project", sourceKey);
    public static string Library(string guid, string sourceKey) => Create("library", guid, sourceKey);
    public static string Type(string libraryId, string sourceKey) => Create("type", libraryId, sourceKey);
    public static string Member(string typeId, string sourceKey) => Create("member", typeId, sourceKey);
    public static string AccessorGroup(string typeId, string name) => Create("accessor", typeId, name);
    public static string Alias(string alias) => Create("alias", alias);
    public static string Unification(string canonicalId) => Create("unification", canonicalId);
    public static string Ambiguity(string kind, string sourceKey) => Create("ambiguity", kind, sourceKey);

    public static string Create(string kind, params string[] components)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(kind);
        if (components.Any(static component => string.IsNullOrWhiteSpace(component)))
            throw new ArgumentException("Logical ID components must not be empty.", nameof(components));

        var value = new StringBuilder("netoffice-data-v2\0").Append(kind.Trim().ToLowerInvariant());
        foreach (var component in components)
            value.Append('\0').Append(component.Normalize(NormalizationForm.FormC));

        var digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value.ToString()))).ToLowerInvariant();
        return $"{kind.Trim().ToLowerInvariant()}-{digest}";
    }
}
