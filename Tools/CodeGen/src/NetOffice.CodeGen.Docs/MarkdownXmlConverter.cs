using System.Xml.Linq;

namespace NetOffice.CodeGen.Docs;

/// <summary>Constrained conversion used only by the pinned VBA profile.</summary>
public static class MarkdownXmlConverter
{
    public static string Convert(string markdown, string title)
    {
        if (markdown is null) throw new ArgumentNullException(nameof(markdown));
        var lines = markdown.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        var summary = new List<string>();
        var remarks = new List<string>();
        var parameters = new List<(string Name, string Text)>();
        var target = summary;
        foreach (var raw in lines.Skip(1))
        {
            var line = raw.Trim();
            if (line.Length == 0) { if (target == summary) target = remarks; continue; }
            if (line.StartsWith("## ", StringComparison.Ordinal)) { target = remarks; continue; }
            if (line.StartsWith("- `", StringComparison.Ordinal) && line.Contains("`:", StringComparison.Ordinal))
            {
                var end = line.IndexOf("`:", StringComparison.Ordinal);
                parameters.Add((line[3..end], Clean(line[(end + 2)..])));
                continue;
            }
            if (!line.StartsWith('#')) target.Add(Clean(line));
        }
        var root = new XElement("member", new XAttribute("name", ""), new XElement("summary", string.Join(" ", summary).Trim()));
        if (parameters.Count > 0) root.Add(parameters.Select(static parameter => new XElement("param", new XAttribute("name", parameter.Name), parameter.Text)));
        if (remarks.Count > 0) root.Add(new XElement("remarks", string.Join(" ", remarks).Trim()));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static string Clean(string value)
        => value.Replace("**", "", StringComparison.Ordinal).Replace("`", "", StringComparison.Ordinal).Trim();
}
