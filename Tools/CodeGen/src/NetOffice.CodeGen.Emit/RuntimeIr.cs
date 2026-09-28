using System;
using System.Collections.Generic;
using System.Linq;

namespace NetOffice.CodeGen.Emit
{
    /// <summary>Runtime API selected for a generated invocation body.</summary>
    public enum WrapperInvocationApi
    {
        Factory,
        Invoker
    }

    public enum WrapperInvocationArgumentPacking
    {
        Flat,
        ObjectArray
    }
    public enum WrapperObjectArrayStyle
    {
        Spaced,
        Compact
    }


    public enum WrapperRawCallCast
    {
        None,
        Object
    }

    public enum WrapperKnownReferenceFactoryStyle
    {
        Generic,
        NonGeneric
    }

    public enum WrapperReturnValueStyle
    {
        Direct,
        Local
    }
    public enum WrapperReturnCastStyle
    {
        Default,
        Explicit,
        As
    }
    public enum WrapperInvokerCallStyle
    {
        Standard,
        Direct
    }

    public enum WrapperStatementTerminatorStyle
    {
        Standard,
        DoubleSemicolon
    }



    /// <summary>COM operation represented by an invocation body.</summary>
    public enum WrapperInvocationKind
    {
        Method,
        PropertyGet,
        PropertySetVariant,
        PropertySet,
        PropertySetValue,
        PropertySetEnum,
        PropertyPutRef,
        LocalForward,
        EventRaise
    }

    /// <summary>How a COM return value is projected into the declared managed return type.</summary>
    public enum WrapperReturnKind
    {
        Void,
        Raw,
        Scalar,
        Enum,
        Struct,
        Variant,
        Reference,
        UntypedReference,
        KnownReference,
        BaseReference,
        EventArgument
    }
    /// <summary>Typed conversion applied by a synthesized COM event sink.</summary>
    public enum WrapperEventArgumentKind
    {
        Raw,
        Scalar,
        Enum,
        KnownReference,
        EventReference
    }

    /// <summary>Projects a native event argument before dispatch and optionally converts its ref/out writeback.</summary>
    public sealed class WrapperEventArgumentConversion
    {
        public WrapperEventArgumentKind Kind { get; set; }
        public string ManagedType { get; set; }
        public string WrapperTypeExpression { get; set; }
        public string LocalName { get; set; }
        public string SourceArgument { get; set; }
        public string ConversionExpression { get; set; }
        public string WriteBackExpression { get; set; }

        internal WrapperEventArgumentConversion Canonicalize()
        {
            var result = new WrapperEventArgumentConversion
            {
                Kind = Kind,
                ManagedType = Normalize(ManagedType),
                WrapperTypeExpression = Normalize(WrapperTypeExpression),
                ConversionExpression = Normalize(ConversionExpression),
                LocalName = Normalize(LocalName),
                SourceArgument = Normalize(SourceArgument),
                WriteBackExpression = Normalize(WriteBackExpression)
            };
            if (result.Kind != WrapperEventArgumentKind.Raw && string.IsNullOrWhiteSpace(result.ManagedType))
                throw new ArgumentException("A converted event argument requires ManagedType.");
            if (result.Kind == WrapperEventArgumentKind.KnownReference && string.IsNullOrWhiteSpace(result.WrapperTypeExpression))
                throw new ArgumentException("A known event reference requires WrapperTypeExpression.");
            if (!string.IsNullOrWhiteSpace(result.ConversionExpression) && result.ConversionExpression.IndexOf("{0}", StringComparison.Ordinal) < 0)
                throw new ArgumentException("ConversionExpression must contain {0}.");
            if (!string.IsNullOrWhiteSpace(result.WriteBackExpression) && result.WriteBackExpression.IndexOf("{0}", StringComparison.Ordinal) < 0)
                throw new ArgumentException("WriteBackExpression must contain {0}.");
            return result;
        }

        private static string Normalize(string value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }


    public enum WrapperConstructorProfile
    {
        Wrapper,
        CoClass
    }

    /// <summary>Exact contract facets layered onto an emitter-owned runtime member body.</summary>
    public sealed class WrapperRuntimeMemberContract
    {
        public IList<string> Attributes { get; set; } = new List<string>();
        public WrapperDocumentation Documentation { get; set; }

        internal WrapperRuntimeMemberContract Canonicalize()
        {
            return new WrapperRuntimeMemberContract
            {
                Attributes = WrapperFile.NormalizeLines(Attributes),
                Documentation = Documentation == null ? null : Documentation.Canonicalize()
            };
        }
    }

    /// <summary>Exact public contract facets for the emitter-owned standard constructor set.</summary>
    public sealed class WrapperConstructorContracts
    {
        public WrapperRuntimeMemberContract ProxyShare { get; set; }
        public WrapperRuntimeMemberContract FactoryParentProxy { get; set; }
        public WrapperRuntimeMemberContract ParentProxy { get; set; }
        public WrapperRuntimeMemberContract FactoryParentProxyType { get; set; }
        public WrapperRuntimeMemberContract ParentProxyType { get; set; }
        public WrapperRuntimeMemberContract ReplacedObject { get; set; }
        public WrapperRuntimeMemberContract Default { get; set; }
        public WrapperRuntimeMemberContract ProgId { get; set; }

        internal WrapperConstructorContracts Canonicalize()
        {
            return new WrapperConstructorContracts
            {
                ProxyShare = Canonicalize(ProxyShare),
                FactoryParentProxy = Canonicalize(FactoryParentProxy),
                ParentProxy = Canonicalize(ParentProxy),
                FactoryParentProxyType = Canonicalize(FactoryParentProxyType),
                ParentProxyType = Canonicalize(ParentProxyType),
                ReplacedObject = Canonicalize(ReplacedObject),
                Default = Canonicalize(Default),
                ProgId = Canonicalize(ProgId)
            };
        }

        private static WrapperRuntimeMemberContract Canonicalize(WrapperRuntimeMemberContract value)
        {
            return value == null ? null : value.Canonicalize();
        }
    }

    /// <summary>Describes the standard NetOffice runtime surface of a wrapper class.</summary>
    public sealed class WrapperRuntime
    {
        public bool EmitInstanceType { get; set; } = true;
        public bool EmitLateBindingApiWrapperType { get; set; } = true;
        public bool EmitStandardConstructors { get; set; } = true;
        public string CoreType { get; set; } = "Core";
        public WrapperConstructorProfile ConstructorProfile { get; set; }
        public string ProgId { get; set; }
        public string ComObjectType { get; set; } = "ICOMObject";
        public string ProxyShareType { get; set; } = "COMProxyShare";
        public string RuntimeType { get; set; } = "Type";
        public string ProxyType { get; set; } = "NetRuntimeSystem.Type";
        public WrapperRuntimeMemberContract InstanceTypeContract { get; set; }
        public WrapperRuntimeMemberContract LateBindingApiWrapperTypeContract { get; set; }
        public WrapperConstructorContracts ConstructorContracts { get; set; }

        internal WrapperRuntime Canonicalize()
        {
            var result = new WrapperRuntime
            {
                EmitInstanceType = EmitInstanceType,
                EmitLateBindingApiWrapperType = EmitLateBindingApiWrapperType,
                EmitStandardConstructors = EmitStandardConstructors,
                ConstructorProfile = ConstructorProfile,
                ProgId = NormalizeOptional(ProgId),
                CoreType = Normalize(CoreType, "Core"),
                ComObjectType = Normalize(ComObjectType, "ICOMObject"),
                ProxyShareType = Normalize(ProxyShareType, "COMProxyShare"),
                RuntimeType = Normalize(RuntimeType, "Type"),
                ProxyType = Normalize(ProxyType, "NetRuntimeSystem.Type"),
                InstanceTypeContract = InstanceTypeContract == null ? null : InstanceTypeContract.Canonicalize(),
                LateBindingApiWrapperTypeContract = LateBindingApiWrapperTypeContract == null ? null : LateBindingApiWrapperTypeContract.Canonicalize(),
                ConstructorContracts = ConstructorContracts == null ? null : ConstructorContracts.Canonicalize()
            };
            if (result.EmitInstanceType && !result.EmitLateBindingApiWrapperType)
                throw new ArgumentException("InstanceType requires LateBindingApiWrapperType.");
            return result;
        }

        private static string Normalize(string value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        private static string NormalizeOptional(string value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    /// <summary>Runtime infrastructure and invocation receivers for a static COM module.</summary>
    public sealed class WrapperModuleRuntime
    {
        public string InstanceType { get; set; } = "global::NetOffice.ICOMObject";
        public string CoreType { get; set; } = "global::NetOffice.Core";
        public string InvokerType { get; set; } = "global::NetOffice.Invoker";
        public string InstanceField { get; set; } = "_instance";
        public string InstanceProperty { get; set; } = "Instance";
        public string FactoryProperty { get; set; } = "Factory";
        public string InvokerProperty { get; set; } = "Invoker";

        internal WrapperModuleRuntime Canonicalize()
        {
            return new WrapperModuleRuntime
            {
                InstanceType = Normalize(InstanceType, "global::NetOffice.ICOMObject"),
                CoreType = Normalize(CoreType, "global::NetOffice.Core"),
                InvokerType = Normalize(InvokerType, "global::NetOffice.Invoker"),
                InstanceField = Normalize(InstanceField, "_instance"),
                InstanceProperty = Normalize(InstanceProperty, "Instance"),
                FactoryProperty = Normalize(FactoryProperty, "Factory"),
                InvokerProperty = Normalize(InvokerProperty, "Invoker")
            };
        }

        private static string Normalize(string value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
    }

    /// <summary>A typed argument passed to Factory/Invoker, optionally copied back after a ref/out call.</summary>
    public sealed class WrapperInvocationArgument
    {
        public string Expression { get; set; }
        public string WriteBackExpression { get; set; }
        public string WriteBackType { get; set; }
        public bool ByRef { get; set; }
        public string InitializationExpression { get; set; }
        public string WriteBackConversion { get; set; }
        public bool IsPropertyValue { get; set; }

        internal WrapperInvocationArgument Canonicalize()
        {
            if (string.IsNullOrWhiteSpace(Expression)) throw new ArgumentException("WrapperInvocationArgument.Expression is required.");
            var result = new WrapperInvocationArgument
            {
                Expression = Expression.Trim(),
                WriteBackExpression = string.IsNullOrWhiteSpace(WriteBackExpression) ? null : WriteBackExpression.Trim(),
                WriteBackType = string.IsNullOrWhiteSpace(WriteBackType) ? "object" : WriteBackType.Trim(),
                ByRef = ByRef,
                IsPropertyValue = IsPropertyValue,
                InitializationExpression = string.IsNullOrWhiteSpace(InitializationExpression) ? null : InitializationExpression.Trim(),
                WriteBackConversion = string.IsNullOrWhiteSpace(WriteBackConversion) ? null : WriteBackConversion.Trim()
            };
            if (!string.IsNullOrWhiteSpace(result.WriteBackConversion) && result.WriteBackConversion.IndexOf("{0}", StringComparison.Ordinal) < 0)
                throw new ArgumentException("WriteBackConversion must contain {0}.");
            if (!string.IsNullOrWhiteSpace(result.InitializationExpression) && !result.ByRef)
                throw new ArgumentException("Only by-ref arguments may define InitializationExpression.");
            return result;
        }
    }

    /// <summary>Canonical, product-independent recipe for a NetOffice runtime invocation body.</summary>
    public sealed class WrapperInvocation
    {
        public WrapperInvocationApi Api { get; set; } = WrapperInvocationApi.Factory;
        public WrapperInvocationKind Kind { get; set; } = WrapperInvocationKind.Method;
        public WrapperReturnKind ReturnKind { get; set; } = WrapperReturnKind.Void;
        public string Target { get; set; }
        public string DispatchName { get; set; }
        public string ReturnType { get; set; }
        public string WrapperTypeExpression { get; set; }
        public string ScalarConversion { get; set; }
        public string EventName { get; set; }
        public string FactoryExpression { get; set; } = "Factory";
        public string InvokerExpression { get; set; } = "Invoker";
        public bool ReleaseArguments { get; set; }
        public WrapperInvocationArgumentPacking ArgumentPacking { get; set; }
        public WrapperObjectArrayStyle ObjectArrayStyle { get; set; }
        public WrapperRawCallCast RawCallCast { get; set; }
        public WrapperKnownReferenceFactoryStyle KnownReferenceFactoryStyle { get; set; }
        public WrapperReturnValueStyle ReturnValueStyle { get; set; }
        public string ReturnLocalName { get; set; } = "newObject";
        public string ReturnLocalType { get; set; }
        public WrapperReturnCastStyle ReturnCastStyle { get; set; } = WrapperReturnCastStyle.As;
        public WrapperInvokerCallStyle InvokerCallStyle { get; set; }
        public WrapperStatementTerminatorStyle StatementTerminatorStyle { get; set; }
        public IList<WrapperInvocationArgument> Arguments { get; set; } = new List<WrapperInvocationArgument>();

        internal WrapperInvocation Canonicalize()
        {
            var result = new WrapperInvocation
            {
                Api = Api,
                Kind = Kind,
                ReturnKind = ReturnKind,
                Target = Normalize(Target, Kind == WrapperInvocationKind.EventRaise ? "EventClass" : Kind == WrapperInvocationKind.LocalForward ? null : "this"),
                DispatchName = Normalize(DispatchName, null),
                ReturnType = Normalize(ReturnType, null),
                WrapperTypeExpression = Normalize(WrapperTypeExpression, null),
                ScalarConversion = Normalize(ScalarConversion, null),
                EventName = Normalize(EventName, null),
                ReleaseArguments = ReleaseArguments,
                ArgumentPacking = ArgumentPacking,
                ObjectArrayStyle = ObjectArrayStyle,
                RawCallCast = RawCallCast,
                KnownReferenceFactoryStyle = KnownReferenceFactoryStyle,
                ReturnValueStyle = ReturnValueStyle,
                ReturnLocalName = Normalize(ReturnLocalName, "newObject"),
                ReturnLocalType = Normalize(ReturnLocalType, Normalize(ReturnType, null)),
                ReturnCastStyle = ReturnCastStyle,
                InvokerCallStyle = InvokerCallStyle,
                StatementTerminatorStyle = StatementTerminatorStyle,
                FactoryExpression = Normalize(FactoryExpression, "Factory"),
                InvokerExpression = Normalize(InvokerExpression, "Invoker"),
                Arguments = (Arguments ?? new List<WrapperInvocationArgument>()).Select(x => x.Canonicalize()).ToList()
            };
            result.Validate();
            return result;
        }

        internal void Validate()
        {
            if (Kind != WrapperInvocationKind.EventRaise && string.IsNullOrWhiteSpace(DispatchName))
                throw new ArgumentException("WrapperInvocation.DispatchName is required.");
            if (Kind == WrapperInvocationKind.EventRaise && string.IsNullOrWhiteSpace(EventName) && string.IsNullOrWhiteSpace(DispatchName))
                throw new ArgumentException("EventRaise requires EventName or DispatchName.");
            if (Kind == WrapperInvocationKind.EventRaise && ReturnKind != WrapperReturnKind.Void)
                throw new ArgumentException("EventRaise cannot return a value.");
            if (Kind == WrapperInvocationKind.PropertyGet && ReturnKind == WrapperReturnKind.Void)
                throw new ArgumentException("PropertyGet requires a return value.");
            if (ReturnKind != WrapperReturnKind.Void && string.IsNullOrWhiteSpace(ReturnType))
                throw new ArgumentException("A non-void invocation requires ReturnType.");
            if (ReturnKind == WrapperReturnKind.KnownReference && string.IsNullOrWhiteSpace(WrapperTypeExpression))
                throw new ArgumentException("KnownReference requires WrapperTypeExpression.");
            if (Api == WrapperInvocationApi.Factory && ReturnKind == WrapperReturnKind.EventArgument)
                throw new ArgumentException("EventArgument returns require the Invoker API.");
            if (Kind != WrapperInvocationKind.EventRaise && Kind != WrapperInvocationKind.LocalForward && Api == WrapperInvocationApi.Factory && Arguments.Any(argument => argument.ByRef))
                throw new ArgumentException("By-ref arguments require the Invoker API.");
            if (InvokerCallStyle == WrapperInvokerCallStyle.Direct && (Api != WrapperInvocationApi.Invoker || Arguments.Count != 0 || ReleaseArguments || (ReturnKind != WrapperReturnKind.Raw && ReturnKind != WrapperReturnKind.Void)))
                throw new ArgumentException("Direct Invoker calls require the Invoker API, zero arguments, no release block, and a raw or void return.");
            if (Kind == WrapperInvocationKind.PropertySet || Kind == WrapperInvocationKind.PropertySetValue || Kind == WrapperInvocationKind.PropertySetVariant || Kind == WrapperInvocationKind.PropertySetEnum || Kind == WrapperInvocationKind.PropertyPutRef)
            {
                if (ReturnKind != WrapperReturnKind.Void) throw new ArgumentException("Property setters cannot return a value.");
                if (Arguments.Count == 0) throw new ArgumentException("Property setters require a value argument.");
                if (Arguments.Count(argument => argument.IsPropertyValue) > 1) throw new ArgumentException("Property setters may mark only one property value argument.");
            }
        }

        private static string Normalize(string value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
    }
}
