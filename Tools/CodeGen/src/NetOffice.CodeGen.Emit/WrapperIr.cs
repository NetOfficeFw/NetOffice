using System;
using System.Collections.Generic;
using System.Linq;

namespace NetOffice.CodeGen.Emit
{
    /// <summary>Canonical input to the source emitter. Lists are copied and canonicalized by Canonicalize.</summary>
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
        public bool Static { get; set; }
        public WrapperRuntime Runtime { get; set; }
        public WrapperModuleRuntime ModuleRuntime { get; set; }
        public WrapperEventBinding EventBinding { get; set; }
        public WrapperProjectInfo ProjectInfo { get; set; }
        public IList<string> Constraints { get; set; } = new List<string>();
        public string TypeParameters { get; set; }
        public string DelegateReturnType { get; set; } = "void";
        public IList<WrapperParameter> DelegateParameters { get; set; } = new List<WrapperParameter>();
        public IList<string> BaseTypes { get; set; } = new List<string>();
        public WrapperDocumentation Documentation { get; set; }
        public IList<string> Attributes { get; set; } = new List<string>();
        public IList<WrapperMember> Members { get; set; } = new List<WrapperMember>();

        public WrapperType Canonicalize()
        {
            var kind = (Kind ?? "class").Trim().ToLowerInvariant();
            var members = (Members ?? new List<WrapperMember>()).Select(x => x.Canonicalize()).ToList();
            var copy = new WrapperType
            {
                Name = (Name ?? string.Empty).Trim(),
                Kind = kind,
                Accessibility = (Accessibility ?? "public").Trim(),
                Partial = Partial,
                Sealed = Sealed,
                Abstract = Abstract,
                TypeParameters = (TypeParameters ?? string.Empty).Trim(),
                Static = Static,
                Runtime = Runtime == null ? null : Runtime.Canonicalize(),
                ModuleRuntime = ModuleRuntime == null ? null : ModuleRuntime.Canonicalize(),
                EventBinding = EventBinding == null ? null : EventBinding.Canonicalize(),
                ProjectInfo = ProjectInfo == null ? null : ProjectInfo.Canonicalize(),
                Constraints = WrapperFile.NormalizeLines(Constraints),
                DelegateReturnType = (DelegateReturnType ?? "void").Trim(),
                DelegateParameters = (DelegateParameters ?? new List<WrapperParameter>()).Select(x => x.Canonicalize()).ToList(),
                BaseTypes = DistinctPreservingOrder(BaseTypes),
                Documentation = Documentation == null ? null : Documentation.Canonicalize(),
                Attributes = WrapperFile.NormalizeLines(Attributes),
                Members = (kind == "enum" ? members.OrderBy(x => x.DeclarationOrder).ThenBy(x => x.SortKey, StringComparer.Ordinal) : members.OrderBy(x => x.SortKey, StringComparer.Ordinal)).ToList()
            };
            if (copy.Kind == "module" && copy.ModuleRuntime == null) copy.ModuleRuntime = new WrapperModuleRuntime().Canonicalize();
            if (copy.ProjectInfo != null && !copy.BaseTypes.Contains(copy.ProjectInfo.InterfaceType, StringComparer.Ordinal))
                copy.BaseTypes.Add(copy.ProjectInfo.InterfaceType);
            copy.Validate();
            return copy;
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Name)) throw new ArgumentException("WrapperType.Name is required.");
            if (Kind != "class" && Kind != "interface" && Kind != "struct" && Kind != "enum" && Kind != "delegate" && Kind != "module" && Kind != "constants") throw new ArgumentException("Unsupported wrapper type kind: " + Kind);
            if (Kind == "enum" && Members.Any(x => x.Kind != "enum-value")) throw new ArgumentException("Enum members must use Kind=enum-value.");
            if (Kind == "delegate" && Members.Count != 0) throw new ArgumentException("Delegates use DelegateParameters rather than members.");
            foreach (var parameter in DelegateParameters) parameter.Validate();
            if (Kind == "interface" && Members.Any(x => x.Kind == "field" || x.Kind == "constant" || x.Kind == "constructor" || x.Kind == "ctor")) throw new ArgumentException("Interfaces cannot contain fields, constants, or constructors.");
            if (Kind == "struct" && Abstract) throw new ArgumentException("Structs cannot be abstract.");
            if (Abstract && Sealed) throw new ArgumentException("A type cannot be both abstract and sealed; use Static for a static class.");
            if (Static && (Abstract || Sealed || BaseTypes.Count != 0 || Runtime != null)) throw new ArgumentException("Static classes cannot specify abstract/sealed modifiers, base types, or runtime metadata.");
            if ((Kind == "module" || Kind == "constants") && (Abstract || Sealed || Runtime != null)) throw new ArgumentException("Modules and constants containers are implicitly static.");
            if ((Kind == "module" || Kind == "constants") && BaseTypes.Count != 0) throw new ArgumentException("Modules and constants containers cannot specify base types.");
            if ((Kind == "module" || Kind == "constants" || Static) && Members.Any(x => x.Kind == "constructor" || x.Kind == "ctor")) throw new ArgumentException("Static containers cannot contain constructors.");
            if (Kind == "constants" && Members.Any(x => x.Kind != "constant")) throw new ArgumentException("Constants containers may contain only constant members.");
            if (Runtime != null && Kind != "class") throw new ArgumentException("Only classes may define NetOffice runtime metadata.");
            if (ModuleRuntime != null && Kind != "module") throw new ArgumentException("Only modules may define module runtime metadata.");
            if (EventBinding != null && Kind != "class") throw new ArgumentException("Only classes may define event-binding runtime metadata.");
            if (ProjectInfo != null && Kind != "class") throw new ArgumentException("Only classes may define ProjectInfo runtime metadata.");
            if (ProjectInfo != null && Members.Count != 0) throw new ArgumentException("ProjectInfo members are emitter-owned and cannot be supplied as regular members.");
            if (Kind == "interface" && Members.Any(x => x.Static)) throw new ArgumentException("C# 7.3 interfaces cannot contain static members.");
            foreach (var member in Members) member.Validate();
        }

        private static IList<string> DistinctPreservingOrder(IEnumerable<string> source)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var value in source ?? Enumerable.Empty<string>())
            {
                var normalized = (value ?? string.Empty).Trim();
                if (normalized.Length != 0 && seen.Add(normalized)) result.Add(normalized);
            }
            return result;
        }
    }
    public enum WrapperEventValidationReleaseMode
    {
        Default,
        Suppress,
        Explicit
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
        public string Initializer { get; set; }
        public string Value { get; set; }
        public string ExplicitInterface { get; set; }
        public string TypeParameters { get; set; }
        public WrapperInvocation Invocation { get; set; }
        public bool Static { get; set; }
        public bool Virtual { get; set; }
        public bool Override { get; set; }
        public bool Abstract { get; set; }
        public bool ReadOnly { get; set; }
        public bool Const { get; set; }
        public bool New { get; set; }
        /// <summary>Preserves contract-owned parameter attributes on a concrete member instead of treating them as native-interface metadata.</summary>
        public bool PreserveParameterAttributes { get; set; }
        public int DeclarationOrder { get; set; }
        public bool UseCustomEventAccessors { get; set; }
        public string EventBackingField { get; set; }
        public string EventValidationKey { get; set; }
        public IList<string> EventValidationReleaseArguments { get; set; }
        public WrapperEventValidationReleaseMode EventValidationReleaseMode { get; set; }
        public bool EventValidationInlineReturn { get; set; }
        public bool EventSinkOnly { get; set; }
        public IList<string> Constraints { get; set; } = new List<string>();
        public IList<string> Attributes { get; set; } = new List<string>();
        public IList<WrapperParameter> Parameters { get; set; } = new List<WrapperParameter>();
        public IList<WrapperAccessor> Accessors { get; set; } = new List<WrapperAccessor>();
        public WrapperDocumentation Documentation { get; set; }
        public string SortKey
        {
            get
            {
                var parameters = string.Join(",", (Parameters ?? new List<WrapperParameter>()).Select(x => (x.Ref ? "ref " : x.Out ? "out " : x.In ? "in " : x.Params ? "params " : string.Empty) + (x.Type ?? string.Empty) + " " + (x.Name ?? string.Empty) + "=" + (x.DefaultValue ?? string.Empty)));
                return (Name ?? string.Empty) + "\u0000" + (Kind ?? string.Empty) + "\u0000" + (Type ?? string.Empty) + "\u0000" + (ExplicitInterface ?? string.Empty) + "\u0000" + (Declaration ?? string.Empty) + "\u0000" + parameters;
            }
        }

        public WrapperMember Canonicalize()
        {
            var kind = (Kind ?? "method").Trim().ToLowerInvariant();
            var invocation = Invocation == null ? null : Invocation.Canonicalize();
            var type = (Type ?? string.Empty).Trim();
            if (kind == "method" && type.Length == 0)
                type = invocation != null && invocation.ReturnKind != WrapperReturnKind.Void ? invocation.ReturnType : "void";
            return new WrapperMember
            {
                Name = (Name ?? string.Empty).Trim(), Kind = kind, Accessibility = (Accessibility ?? "public").Trim(), Type = type,
                Declaration = Normalize(Declaration), Body = Normalize(Body), ExpressionBody = Normalize(ExpressionBody), Initializer = Normalize(Initializer), Value = Normalize(Value),
                ExplicitInterface = Normalize(ExplicitInterface), TypeParameters = Normalize(TypeParameters), Invocation = invocation,
                Static = Static, Virtual = Virtual, Override = Override, Abstract = Abstract, ReadOnly = ReadOnly, Const = Const, New = New,
                PreserveParameterAttributes = PreserveParameterAttributes,
                UseCustomEventAccessors = UseCustomEventAccessors, EventBackingField = Normalize(EventBackingField),
                DeclarationOrder = DeclarationOrder,
                Constraints = WrapperFile.NormalizeLines(Constraints), Attributes = WrapperFile.NormalizeLines(Attributes), Parameters = (Parameters ?? new List<WrapperParameter>()).Select(x => x.Canonicalize()).ToList(),
                EventValidationKey = Normalize(EventValidationKey),
                EventValidationReleaseArguments = EventValidationReleaseArguments == null ? null : WrapperFile.NormalizeLines(EventValidationReleaseArguments),
                EventValidationReleaseMode = EventValidationReleaseMode,
                EventValidationInlineReturn = EventValidationInlineReturn,
                EventSinkOnly = EventSinkOnly,
                Accessors = (Accessors ?? new List<WrapperAccessor>()).Select(x => x.Canonicalize()).OrderBy(AccessorOrder).ThenBy(x => x.Kind, StringComparer.Ordinal).ToList(), Documentation = Documentation == null ? null : Documentation.Canonicalize()
            };
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Name) && string.IsNullOrWhiteSpace(Declaration)) throw new ArgumentException("WrapperMember.Name or Declaration is required.");
            if (string.IsNullOrWhiteSpace(Kind)) throw new ArgumentException("WrapperMember.Kind is required.");
            if (Kind == "method" && string.IsNullOrWhiteSpace(Declaration) && Parameters.Any(x => string.IsNullOrWhiteSpace(x.Name))) throw new ArgumentException("Method parameter names are required.");
            var kind = Kind.ToLowerInvariant();
            if (Accessors.Count != 0 && kind != "property" && kind != "indexer" && kind != "event") throw new ArgumentException("Only properties, indexers, and events may contain accessors.");
            if (kind == "event" && Accessors.Any(x => x.Kind != "add" && x.Kind != "remove")) throw new ArgumentException("Events may contain only add/remove accessors.");
            if ((kind == "property" || kind == "indexer") && Accessors.Any(x => x.Kind != "get" && x.Kind != "set")) throw new ArgumentException("Properties may contain only get/set accessors.");
            if (Accessors.GroupBy(x => x.Kind, StringComparer.OrdinalIgnoreCase).Any(x => x.Count() > 1)) throw new ArgumentException("A member cannot contain duplicate accessors.");
            if (kind == "event" && Accessors.Count != 0 && (Accessors.Count != 2 || !Accessors.Any(x => x.Kind == "add") || !Accessors.Any(x => x.Kind == "remove"))) throw new ArgumentException("Custom events require one add and one remove accessor.");
            if (UseCustomEventAccessors && kind != "event") throw new ArgumentException("Only events may use custom event accessors.");
            if (UseCustomEventAccessors && string.IsNullOrWhiteSpace(EventBackingField)) throw new ArgumentException("Custom event accessors require EventBackingField.");
            if (UseCustomEventAccessors && Accessors.Count != 0) throw new ArgumentException("Custom event metadata cannot be combined with explicit accessors.");
            if (Invocation != null && (!string.IsNullOrWhiteSpace(Body) || !string.IsNullOrWhiteSpace(ExpressionBody))) throw new ArgumentException("A member cannot specify both Invocation and a body.");
            if (Invocation != null && kind != "method" && kind != "property" && kind != "indexer") throw new ArgumentException("Only methods, properties, and indexers may define a member invocation.");
            if (Invocation != null && (kind == "property" || kind == "indexer") && Accessors.Count != 0) throw new ArgumentException("A property invocation cannot be combined with explicit accessors.");
            if (EventValidationReleaseMode == WrapperEventValidationReleaseMode.Explicit && EventValidationReleaseArguments == null) throw new ArgumentException("Explicit event validation release requires an argument list.");
            if (EventValidationReleaseMode == WrapperEventValidationReleaseMode.Suppress && EventValidationReleaseArguments != null && EventValidationReleaseArguments.Count != 0) throw new ArgumentException("Suppressed event validation release cannot define arguments.");
            if (EventValidationInlineReturn && Type != "void") throw new ArgumentException("Inline event validation return requires a void member.");
            if (EventValidationInlineReturn && EventValidationReleaseMode == WrapperEventValidationReleaseMode.Suppress) throw new ArgumentException("Inline event validation return requires a release call.");
            if (EventValidationInlineReturn && EventValidationReleaseMode == WrapperEventValidationReleaseMode.Default && Parameters.Count == 0) throw new ArgumentException("Inline event validation return requires release arguments.");
            if (EventSinkOnly && kind != "method") throw new ArgumentException("Only methods may be emitted exclusively on an event sink.");
            if ((Const || kind == "constant" || kind == "enum-value") && string.IsNullOrWhiteSpace(Value)) throw new ArgumentException(kind + " requires Value.");
            var polymorphicModifiers = (Virtual ? 1 : 0) + (Override ? 1 : 0) + (Abstract ? 1 : 0);
            if (polymorphicModifiers > 1) throw new ArgumentException("A member can have only one virtual, override, or abstract modifier.");
            if (Static && polymorphicModifiers != 0) throw new ArgumentException("Static members cannot be virtual, override, or abstract.");
            if ((Const || kind == "constant") && kind != "field" && kind != "constant") throw new ArgumentException("Only fields may be const.");
            if (ReadOnly && kind != "field") throw new ArgumentException("Only fields may be readonly.");
            if (Abstract && (!string.IsNullOrWhiteSpace(Body) || !string.IsNullOrWhiteSpace(ExpressionBody) || Invocation != null)) throw new ArgumentException("Abstract members cannot define a body.");
            foreach (var parameter in Parameters) parameter.Validate();
            foreach (var accessor in Accessors) accessor.Validate();
        }

        private static string Normalize(string value) { return string.IsNullOrWhiteSpace(value) ? null : value.Replace("\r\n", "\n").Replace('\r', '\n').Trim(); }
        private static int AccessorOrder(WrapperAccessor accessor)
        {
            if (accessor.Kind == "get" || accessor.Kind == "add") return 0;
            return 1;
        }
    }

    /// <summary>A property or indexer accessor in the canonical wrapper IR.</summary>
    public sealed class WrapperAccessor
    {
        public string Kind { get; set; }
        public string Accessibility { get; set; }
        public string Body { get; set; }
        public string ExpressionBody { get; set; }
        public WrapperInvocation Invocation { get; set; }
        public IList<string> Attributes { get; set; } = new List<string>();
        public WrapperDocumentation Documentation { get; set; }

        public WrapperAccessor Canonicalize()
        {
            return new WrapperAccessor
            {
                Kind = (Kind ?? string.Empty).Trim().ToLowerInvariant(),
                Accessibility = (Accessibility ?? string.Empty).Trim(),
                Body = Normalize(Body),
                ExpressionBody = Normalize(ExpressionBody),
                Invocation = Invocation == null ? null : Invocation.Canonicalize(),
                Attributes = WrapperFile.NormalizeLines(Attributes),
                Documentation = Documentation == null ? null : Documentation.Canonicalize()
            };
        }

        public void Validate()
        {
            if (Kind != "get" && Kind != "set" && Kind != "add" && Kind != "remove" && Kind != "init") throw new ArgumentException("Unsupported accessor kind: " + Kind);
            if (Kind == "init") throw new ArgumentException("C# 7.3 does not support init accessors.");
            if (Invocation != null && (!string.IsNullOrWhiteSpace(Body) || !string.IsNullOrWhiteSpace(ExpressionBody))) throw new ArgumentException("An accessor cannot specify both Invocation and a body.");
            if (Invocation != null && Kind == "get" && Invocation.Kind != WrapperInvocationKind.PropertyGet && Invocation.Kind != WrapperInvocationKind.Method) throw new ArgumentException("A get accessor requires a PropertyGet or Method invocation.");
            if (Invocation != null && Kind == "get" && Invocation.ReturnKind == WrapperReturnKind.Void) throw new ArgumentException("A get accessor invocation must return a value.");
            if (Invocation != null && Kind == "set" && Invocation.Kind != WrapperInvocationKind.PropertySet && Invocation.Kind != WrapperInvocationKind.PropertySetValue && Invocation.Kind != WrapperInvocationKind.PropertySetVariant && Invocation.Kind != WrapperInvocationKind.PropertySetEnum && Invocation.Kind != WrapperInvocationKind.PropertyPutRef) throw new ArgumentException("A set accessor requires a property-set invocation.");
            if (Invocation != null && (Kind == "add" || Kind == "remove")) throw new ArgumentException("Event accessors use an explicit body.");
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
        public bool Params { get; set; }
        public IList<string> Attributes { get; set; } = new List<string>();
        public WrapperEventArgumentConversion EventConversion { get; set; }
        public WrapperDocumentation Documentation { get; set; }
        public WrapperParameter Canonicalize()
        {
            var defaultValue = Normalize(DefaultValue);
            if (Ref || Out) defaultValue = null;
            return new WrapperParameter
            {
                Name = (Name ?? string.Empty).Trim(), Type = (Type ?? "object").Trim(), DefaultValue = defaultValue,
                Ref = Ref, Out = Out, In = In, Params = Params, Attributes = WrapperFile.NormalizeLines(Attributes),
                EventConversion = EventConversion == null ? null : EventConversion.Canonicalize(),
                Documentation = Documentation == null ? null : Documentation.Canonicalize()
            };
        }
        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Name)) throw new ArgumentException("WrapperParameter.Name is required.");
            if (string.IsNullOrWhiteSpace(Type)) throw new ArgumentException("WrapperParameter.Type is required.");
            var modifiers = (Ref ? 1 : 0) + (Out ? 1 : 0) + (In ? 1 : 0) + (Params ? 1 : 0);
            if (modifiers > 1) throw new ArgumentException("A parameter can have only one ref, out, in, or params modifier.");
            if (Params && !string.IsNullOrWhiteSpace(DefaultValue)) throw new ArgumentException("params parameters cannot have default values.");
        }
        private static string Normalize(string value) { return string.IsNullOrWhiteSpace(value) ? null : value.Replace("\r\n", "\n").Replace('\r', '\n').Trim(); }
    }

    public sealed class WrapperDocumentation
    {
        /// <summary>Canonical XML-doc fragment without triple-slash prefixes or member indentation.</summary>
        public string RawXml { get; set; }
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
            return new WrapperDocumentation { RawXml = Normalize(RawXml), Summary = Normalize(Summary), Remarks = Normalize(Remarks), Returns = Normalize(Returns), Value = Normalize(Value), Examples = WrapperFile.NormalizeLines(Examples), Parameters = Sorted(Parameters), TypeParameters = Sorted(TypeParameters), Exceptions = WrapperFile.NormalizeLines(Exceptions) };
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
