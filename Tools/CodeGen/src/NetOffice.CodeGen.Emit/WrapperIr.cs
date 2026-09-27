using System;
using System.Collections.Generic;
using System.Linq;

namespace NetOffice.CodeGen.Emit
{
    /// <summary>Canonical input to the source emitter. Lists are copied and sorted by Canonicalize.</summary>
    public sealed class WrapperFile
    {
        public string RelativePath { get; set; }
        public string Namespace { get; set; }
        public IList<string> Usings { get; set; } = new List<string>();
        public IList<string> HeaderComments { get; set; } = new List<string>();
        public IList<WrapperType> Types { get; set; } = new List<WrapperType>();

        public WrapperFile Canonicalize()
        {
            var result = new WrapperFile
            {
                RelativePath = CanonicalPath(RelativePath),
                Namespace = (Namespace ?? string.Empty).Trim(),
                Usings = DistinctSorted(Usings, StringComparer.Ordinal),
                HeaderComments = NormalizeLines(HeaderComments),
                Types = (Types ?? new List<WrapperType>()).Select(x => x.Canonicalize()).OrderBy(x => x.Name, StringComparer.Ordinal).ThenBy(x => x.Kind, StringComparer.Ordinal).ToList()
            };
            result.Validate();
            return result;
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(RelativePath)) throw new ArgumentException("WrapperFile.RelativePath is required.");
            if (RelativePath.StartsWith("/", StringComparison.Ordinal) || RelativePath.Contains("..", StringComparison.Ordinal)) throw new ArgumentException("RelativePath must not escape the output root.");
            if (Types == null || Types.Count == 0) throw new ArgumentException("A wrapper file must contain at least one type.");
            if (Types.GroupBy(x => x.Name, StringComparer.Ordinal).Any(x => x.Count() > 1)) throw new ArgumentException("A wrapper file cannot contain duplicate type names.");
            foreach (var type in Types) type.Validate();
        }

        internal static string CanonicalPath(string path)
        {
            if (path == null) return string.Empty;
            return path.Replace('\\', '/').TrimStart('/');
        }

        internal static IList<string> DistinctSorted(IEnumerable<string> source, StringComparer comparer)
        {
            return (source ?? Enumerable.Empty<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).Distinct(comparer).OrderBy(x => x, comparer).ToList();
        }

        internal static IList<string> NormalizeLines(IEnumerable<string> source)
        {
            return (source ?? Enumerable.Empty<string>()).SelectMany(x => (x ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')).Select(x => x.TrimEnd()).Where(x => x.Length != 0).ToList();
        }
    }

    public sealed class WrapperType
    {
        public string Name { get; set; }
        public string Kind { get; set; } = "class";
        public string Accessibility { get; set; } = "public";
        public bool Partial { get; set; }
        public bool Sealed { get; set; }
        public bool Abstract { get; set; }
        public string TypeParameters { get; set; }
        public string DelegateReturnType { get; set; } = "void";
        public IList<WrapperParameter> DelegateParameters { get; set; } = new List<WrapperParameter>();
        public IList<string> BaseTypes { get; set; } = new List<string>();
        public WrapperDocumentation Documentation { get; set; }
        public IList<string> Attributes { get; set; } = new List<string>();
        public IList<WrapperMember> Members { get; set; } = new List<WrapperMember>();

        public WrapperType Canonicalize()
        {
            var copy = new WrapperType
            {
                Name = (Name ?? string.Empty).Trim(),
                Kind = (Kind ?? "class").Trim().ToLowerInvariant(),
                Accessibility = (Accessibility ?? "public").Trim(),
                Partial = Partial,
                Sealed = Sealed,
                Abstract = Abstract,
                TypeParameters = (TypeParameters ?? string.Empty).Trim(),
                DelegateReturnType = (DelegateReturnType ?? "void").Trim(),
                DelegateParameters = (DelegateParameters ?? new List<WrapperParameter>()).Select(x => x.Canonicalize()).ToList(),
                BaseTypes = WrapperFile.DistinctSorted(BaseTypes, StringComparer.Ordinal),
                Documentation = Documentation == null ? null : Documentation.Canonicalize(),
                Attributes = WrapperFile.NormalizeLines(Attributes),
                Members = (Members ?? new List<WrapperMember>()).Select(x => x.Canonicalize()).OrderBy(x => x.SortKey, StringComparer.Ordinal).ToList()
            };
            copy.Validate();
            return copy;
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Name)) throw new ArgumentException("WrapperType.Name is required.");
            if (Kind != "class" && Kind != "interface" && Kind != "struct" && Kind != "enum" && Kind != "delegate") throw new ArgumentException("Unsupported wrapper type kind: " + Kind);
            if (Kind == "enum" && Members.Any(x => x.Kind != "enum-value")) throw new ArgumentException("Enum members must use Kind=enum-value.");
            if (Kind == "delegate" && Members.Count != 0) throw new ArgumentException("Delegates use DelegateParameters rather than members.");
            if (Kind == "interface" && Members.Any(x => x.Kind == "field" || x.Kind == "constructor" || x.Kind == "ctor")) throw new ArgumentException("Interfaces cannot contain fields or constructors.");
            foreach (var member in Members) member.Validate();
        }
    }

    public sealed class WrapperMember
    {
        public string Name { get; set; }
        public string Kind { get; set; } = "method";
        public string Accessibility { get; set; } = "public";
        public string Type { get; set; } = "void";
        public string Declaration { get; set; }
        public string Body { get; set; }
        public string ExpressionBody { get; set; }
        public bool Static { get; set; }
        public bool Virtual { get; set; }
        public bool Override { get; set; }
        public bool Abstract { get; set; }
        public bool ReadOnly { get; set; }
        public IList<string> Attributes { get; set; } = new List<string>();
        public IList<WrapperParameter> Parameters { get; set; } = new List<WrapperParameter>();
        public WrapperDocumentation Documentation { get; set; }
        public string SortKey { get { return (Name ?? string.Empty) + "\u0000" + (Kind ?? string.Empty) + "\u0000" + (Declaration ?? string.Empty); } }

        public WrapperMember Canonicalize()
        {
            return new WrapperMember
            {
                Name = (Name ?? string.Empty).Trim(), Kind = (Kind ?? "method").Trim().ToLowerInvariant(), Accessibility = (Accessibility ?? "public").Trim(), Type = (Type ?? "void").Trim(),
                Declaration = Normalize(Declaration), Body = Normalize(Body), ExpressionBody = Normalize(ExpressionBody), Static = Static, Virtual = Virtual, Override = Override, Abstract = Abstract, ReadOnly = ReadOnly,
                Attributes = WrapperFile.NormalizeLines(Attributes), Parameters = (Parameters ?? new List<WrapperParameter>()).Select(x => x.Canonicalize()).ToList(), Documentation = Documentation == null ? null : Documentation.Canonicalize()
            };
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Name) && string.IsNullOrWhiteSpace(Declaration)) throw new ArgumentException("WrapperMember.Name or Declaration is required.");
            if (string.IsNullOrWhiteSpace(Kind)) throw new ArgumentException("WrapperMember.Kind is required.");
            if (Kind == "method" && string.IsNullOrWhiteSpace(Declaration) && Parameters.Any(x => string.IsNullOrWhiteSpace(x.Name))) throw new ArgumentException("Method parameter names are required.");
        }

        private static string Normalize(string value) { return string.IsNullOrWhiteSpace(value) ? null : value.Replace("\r\n", "\n").Replace('\r', '\n').Trim(); }
    }

    public sealed class WrapperParameter
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string DefaultValue { get; set; }
        public bool Ref { get; set; }
        public bool Out { get; set; }
        public bool In { get; set; }
        public WrapperDocumentation Documentation { get; set; }
        public WrapperParameter Canonicalize() { return new WrapperParameter { Name = (Name ?? string.Empty).Trim(), Type = (Type ?? "object").Trim(), DefaultValue = Normalize(DefaultValue), Ref = Ref, Out = Out, In = In, Documentation = Documentation == null ? null : Documentation.Canonicalize() }; }
        private static string Normalize(string value) { return string.IsNullOrWhiteSpace(value) ? null : value.Replace("\r\n", "\n").Replace('\r', '\n').Trim(); }
    }

    public sealed class WrapperDocumentation
    {
        public string Summary { get; set; }
        public string Remarks { get; set; }
        public string Returns { get; set; }
        public string Value { get; set; }
        public IList<string> Examples { get; set; } = new List<string>();
        public IDictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
        public IDictionary<string, string> TypeParameters { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
        public IList<string> Exceptions { get; set; } = new List<string>();

        public WrapperDocumentation Canonicalize()
        {
            return new WrapperDocumentation { Summary = Normalize(Summary), Remarks = Normalize(Remarks), Returns = Normalize(Returns), Value = Normalize(Value), Examples = WrapperFile.NormalizeLines(Examples), Parameters = Sorted(Parameters), TypeParameters = Sorted(TypeParameters), Exceptions = WrapperFile.NormalizeLines(Exceptions) };
        }

        private static IDictionary<string, string> Sorted(IDictionary<string, string> source)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var pair in (source ?? new Dictionary<string, string>()).OrderBy(x => x.Key, StringComparer.Ordinal)) result[pair.Key.Trim()] = Normalize(pair.Value);
            return result;
        }
        private static string Normalize(string value) { return string.IsNullOrWhiteSpace(value) ? null : value.Replace("\r\n", "\n").Replace('\r', '\n').Trim(); }
    }
}
