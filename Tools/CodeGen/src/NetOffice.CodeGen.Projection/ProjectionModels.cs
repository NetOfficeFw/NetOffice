using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Projection;

public sealed record ProjectionOptions
{
    public string? ExpectedDataDigest { get; init; }
    public string? ExpectedContractDigest { get; init; }
    public string? ExpectedPolicyDigest { get; init; }
    public bool RejectUnresolvedAmbiguities { get; init; } = true;
    public bool IncludeTraces { get; init; } = true;
}

public sealed record WrapperContract
{
    public string SchemaVersion { get; init; } = "";
    public string ContractKind { get; init; } = "";
    public ContractGenerator Generator { get; init; } = new();
    public ContractSource Source { get; init; } = new();
    public IReadOnlyList<ContractType> Types { get; init; } = Array.Empty<ContractType>();
    public IReadOnlyList<ContractUnknown> Unknowns { get; init; } = Array.Empty<ContractUnknown>();
    public IReadOnlyList<ContractAmbiguity> Ambiguities { get; init; } = Array.Empty<ContractAmbiguity>();

    public static WrapperContract Parse(string json)
        => JsonSerializer.Deserialize<WrapperContract>(json, ProjectionPolicy.JsonOptions)
           ?? throw new ProjectionValidationException(new[] { new ProjectionIssue("contract.empty", "Contract JSON is empty.", "") });
    public static WrapperContract Read(string path) => Parse(File.ReadAllText(path));
}

public sealed record ContractGenerator { public string Name { get; init; } = ""; public string Version { get; init; } = ""; }
public sealed record ContractSource { public string Root { get; init; } = ""; public string Api { get; init; } = ""; public IReadOnlyList<string> SourceRoots { get; init; } = Array.Empty<string>(); public IReadOnlyList<ContractFile> Files { get; init; } = Array.Empty<ContractFile>(); }
public sealed record ContractFile { public string Path { get; init; } = ""; public string Sha256 { get; init; } = ""; public int LineCount { get; init; } }
public sealed record ContractType
{
    public string LogicalId { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "class";
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? BaseType { get; init; }
    public IReadOnlyList<string> Interfaces { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public string Source { get; init; } = "";
    public int Line { get; init; }
    public string Signature { get; init; } = "";
    public IReadOnlyList<ContractMember> Members { get; init; } = Array.Empty<ContractMember>();
}
public sealed record ContractMember
{
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? ReturnType { get; init; }
    public string? Parameters { get; init; }
    public string Signature { get; init; } = "";
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> InvocationText { get; init; } = Array.Empty<string>();
    public string Source { get; init; } = "";
    public int Line { get; init; }
}
public sealed record ContractUnknown { public string Code { get; init; } = ""; public string Message { get; init; } = ""; public string Path { get; init; } = ""; public int? Line { get; init; } public string? LogicalId { get; init; } }
public sealed record ContractAmbiguity { public string Code { get; init; } = ""; public string Message { get; init; } = ""; public IReadOnlyList<string> Candidates { get; init; } = Array.Empty<string>(); }

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionEntityKind { DispatchInterface, Interface, EventInterface, CoClass, Enum, Module, Constants, Record, TypeDef, Utility }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionFileCategory { DispatchInterfaces, Interfaces, Events, Classes, Enums, Modules, Constants, Records, TypeDefs, Utils }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionEmissionDisposition { Emit, Companion, SuppressedByContract }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionEventConversion { Raw, Scalar, Enum, KnownReference, EventReference }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionInvocationApi { Factory, Invoker, Local, Runtime }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionInvocationOperation { Method, PropertyGet, PropertySet, PropertySetVariant, PropertySetEnum, PropertyPutRef, LocalForward, ActiveInstance, ActiveInstances, Event, EventRaise, Field, Constructor, None }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionInvocationCallKind { Auto, MethodGet, PropertyGet, PropertySet, ValuePropertySet, VariantPropertySet, EnumPropertySet, ReferencePropertySet }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionRuntimeMemberKind { None, Clone, FromProxyService, Dispose, DisposeWithEventBinding }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionArgumentPacking { Flat, ObjectArray }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionObjectArrayStyle { Compact, Spaced }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionReturnCastStyle { Default, Explicit, As }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionRawCallCast { None, Object }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionKnownReferenceFactoryStyle { Generic, NonGeneric }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionReturnValueStyle { Direct, Local }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionInvokerCallStyle { ParamsArray, Direct }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionMemberEmissionDisposition { Emit, Inherited, SupersededByIndexer, SupersededByProperty, SupersededByConstructorPlan, SuppressedByContract }
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ProjectionReturnConversion { None, Scalar, String, Enum, Value, Struct, Variant, KnownReference, Reference, UntypedReference, BaseReference, EventArgument, Native, Array }

public sealed record ProjectionResult
{
    public string SchemaVersion { get; init; } = "projection/v2";
    public string DataSchemaVersion { get; init; } = "";
    public string DataDigest { get; init; } = "";
    public string ContractSchemaVersion { get; init; } = "";
    public string ContractDigest { get; init; } = "";
    public string PolicyDigest { get; init; } = "";
    public IReadOnlyList<WrapperFile> Files { get; init; } = Array.Empty<WrapperFile>();
    public IReadOnlyList<ProjectionTrace> Explain { get; init; } = Array.Empty<ProjectionTrace>();
    public ProjectionCoverage Coverage { get; init; } = new();

    public string ToJson(bool indented = false) => JsonSerializer.Serialize(this, new JsonSerializerOptions(ProjectionPolicy.JsonOptions) { WriteIndented = indented });
}

public sealed record ProjectionCoverage
{
    public int DataTypes { get; init; }
    public int ProjectedTypes { get; init; }
    public int DataMembers { get; init; }
    public int ProjectedDataMembers { get; init; }
    public int DataValues { get; init; }
    public int ProjectedValues { get; init; }
}

public sealed record WrapperFile
{
    public string Path { get; init; } = "";
    public string Namespace { get; init; } = "";
    public ProjectionFileCategory Category { get; init; }
    public IReadOnlyList<ProjectedType> Types { get; init; } = Array.Empty<ProjectedType>();
}
public sealed record ProjectedType
{
    public string CanonicalLogicalId { get; init; } = "";
    public string? ContractPartLogicalId { get; init; }
    public bool IsPrimaryPart { get; init; } = true;
    public string? ContractPartSource { get; init; }
    public string LogicalId { get; init; } = "";
    public string DataLogicalId { get; init; } = "";
    public string LibraryLogicalId { get; init; } = "";
    public string Product { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string Name { get; init; } = "";
    public string CSharpName { get; init; } = "";
    public string Kind { get; init; } = "";
    public ProjectionEntityKind EntityKind { get; init; }
    public ProjectionEmissionDisposition EmissionDisposition { get; init; } = ProjectionEmissionDisposition.Emit;
    public string? CompanionPath { get; init; }
    public ProjectionFileCategory FileCategory { get; init; }
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public string? DeclaredGuid { get; init; }
    public int? TypeLibType { get; init; }
    public bool IsHidden { get; init; }
    public bool IsEarlyBind { get; init; }
    public bool AutomaticQuit { get; init; }
    public bool IsApplicationObject { get; init; }
    public string? BaseType { get; init; }
    public IReadOnlyList<string> Interfaces { get; init; } = Array.Empty<string>();
    public string? AliasTarget { get; init; }
    public string? ProgId { get; init; }
    public EventSinkPlan? EventSink { get; init; }
    public IReadOnlyList<ProjectedEventBinding> EventBindings { get; init; } = Array.Empty<ProjectedEventBinding>();
    public ProjectInfoPlan? ProjectInfo { get; init; }
    public IReadOnlyList<ProjectedAuxiliaryType> AuxiliaryTypes { get; init; } = Array.Empty<ProjectedAuxiliaryType>();
    public IReadOnlyList<string> EffectiveBaseTypes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> DuplicateGroups { get; init; } = Array.Empty<string>();
    public string Signature { get; init; } = "";
    public IReadOnlyList<VersionSupport> SupportVersions { get; init; } = Array.Empty<VersionSupport>();
    public IReadOnlyList<string> Capabilities { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> RuntimeRequirements { get; init; } = Array.Empty<string>();
    public IReadOnlyList<ConstructorPlan> Constructors { get; init; } = Array.Empty<ConstructorPlan>();
    public IReadOnlyList<ProjectedContractMember> ExactContractMembers { get; init; } = Array.Empty<ProjectedContractMember>();
    public string DocsBindingKey { get; init; } = "";
    public string Source { get; init; } = "";
    public int Line { get; init; }
    public IReadOnlyList<ProjectedMember> Members { get; init; } = Array.Empty<ProjectedMember>();
}
public sealed record ConstructorPlan
{
    public string Kind { get; init; } = "proxy";
    public string LogicalId { get; init; } = "";
    public bool IsContractOverlay { get; init; }
    public string Signature { get; init; } = "";
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<ProjectedParameter> Parameters { get; init; } = Array.Empty<ProjectedParameter>();
    public string BaseCall { get; init; } = "";
    public IReadOnlyList<string> RuntimeRequirements { get; init; } = Array.Empty<string>();
    public string DocsBindingKey { get; init; } = "";
    public string Source { get; init; } = "";
    public int Line { get; init; }
}
public sealed record ProjectedContractMember
{
    public string LogicalId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Accessibility { get; init; } = "";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? ReturnType { get; init; }
    public string? Parameters { get; init; }
    public IReadOnlyList<ProjectedParameter> ParameterList { get; init; } = Array.Empty<ProjectedParameter>();
    public string Signature { get; init; } = "";
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<VersionSupport> SupportVersions { get; init; } = Array.Empty<VersionSupport>();
    public string DocsBindingKey { get; init; } = "";
    public string Source { get; init; } = "";
    public int Line { get; init; }
}
public sealed record ProjectedEventBinding
{
    public string LogicalId { get; init; } = "";
    public string SinkHelperType { get; init; } = "";
    public string FieldName { get; init; } = "";
    public string? EventInterfaceLogicalId { get; init; }
    public string? InterfaceId { get; init; }
}
public sealed record ProjectInfoPlan
{
    public string AssemblyNamespace { get; init; } = "";
    public IReadOnlyList<string> ComponentGuids { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Dependencies { get; init; } = Array.Empty<string>();
}

public sealed record EventSinkPlan
{
    public string Name { get; init; } = "";
    public string InterfaceName { get; init; } = "";
    public string? InterfaceId { get; init; }
}
public sealed record ProjectedAuxiliaryType
{
    public string LogicalId { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? BaseType { get; init; }
    public IReadOnlyList<string> Interfaces { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public string Signature { get; init; } = "";
    public string Source { get; init; } = "";
    public int Line { get; init; }
    public string? DelegateReturnType { get; init; }
    public IReadOnlyList<ProjectedParameter> DelegateParameters { get; init; } = Array.Empty<ProjectedParameter>();
    public IReadOnlyList<ConstructorPlan> Constructors { get; init; } = Array.Empty<ConstructorPlan>();
}
public sealed record ProjectedParameter
{
    public string Name { get; init; } = "";
    public string CSharpName { get; init; } = "";
    public string Type { get; init; } = "object";
    public string RefKind { get; init; } = "value";
    public bool IsOptional { get; init; }
    public bool HasDefaultValue { get; init; }
    public string? DefaultValue { get; init; }
    public bool EmitDefaultValue { get; init; }
    public int Position { get; init; }
    public ProjectedTypeReference? TypeReference { get; init; }
    public bool IsContractOverlay { get; init; }
    public string? ContractDeclaration { get; init; }
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public bool PreserveContractAttributes { get; init; }
    public EventSinkArgumentPlan? SinkArgument { get; init; }
}
public sealed record EventSinkArgumentPlan
{
    public ProjectionEventConversion Conversion { get; init; }
    public string ManagedType { get; init; } = "object";
    public string? WrapperTypeExpression { get; init; }
    public string? ConversionExpression { get; init; }
    public string? WriteBackExpression { get; init; }
    public string? LocalName { get; init; }
    public string? SourceArgument { get; init; }
    public bool IsContractDeclared { get; init; }
}
public sealed record ProjectedTypeReference
{
    public string Name { get; init; } = "";
    public string? TypeKind { get; init; }
    public string? TargetTypeId { get; init; }
    public string? QualifiedName { get; init; }
    public string? VarType { get; init; }
    public string? MarshalAs { get; init; }
    public string? TypeKey { get; init; }
    public string? ProjectKey { get; init; }
    public string? LibraryKey { get; init; }
    public bool IsComProxy { get; init; }
    public bool IsExternal { get; init; }
    public bool IsEnum { get; init; }
    public bool IsArray { get; init; }
    public bool IsNative { get; init; }
}
public sealed record ProjectedMember
{
    public string LogicalId { get; init; } = "";
    public string DataLogicalId { get; init; } = "";
    public string EmissionOwnerLogicalId { get; init; } = "";
    public string Name { get; init; } = "";
    public string CSharpName { get; init; } = "";
    public ProjectionMemberEmissionDisposition EmissionDisposition { get; init; } = ProjectionMemberEmissionDisposition.Emit;
    public string Kind { get; init; } = "";
    public string Accessibility { get; init; } = "";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? ReturnType { get; init; }
    public string? Parameters { get; init; }
    public IReadOnlyList<ProjectedParameter> ParameterList { get; init; } = Array.Empty<ProjectedParameter>();
    public string Signature { get; init; } = "";
    public string OverloadGroup { get; init; } = "";
    public int OverloadOrdinal { get; init; }
    public ProjectionRuntimeMemberKind RuntimeMemberKind { get; init; }
    public string? DuplicateOf { get; init; }
    public string? AccessorGroupId { get; init; }
    public InvocationPlan Invocation { get; init; } = new();
    public IReadOnlyList<InvocationPlan> AccessorInvocations { get; init; } = Array.Empty<InvocationPlan>();
    public bool IsContractDerived { get; init; }
    public IReadOnlyList<VersionSupport> SupportVersions { get; init; } = Array.Empty<VersionSupport>();
    public IReadOnlyList<string> Capabilities { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> RuntimeRequirements { get; init; } = Array.Empty<string>();
    public bool IsHidden { get; init; }
    public bool AnalyzeReturn { get; init; }
    public bool IsComProxy { get; init; }
    public string DocsBindingKey { get; init; } = "";
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public string? ConstantValue { get; init; }
    public string? ConstantExpression { get; init; }
    public string? ValueType { get; init; }
    public IReadOnlyList<string> Accessors { get; init; } = Array.Empty<string>();
    public bool UseCustomEventAccessors { get; init; }
    public string? EventBackingField { get; init; }
    public string? EventValidationKey { get; init; }
    public IReadOnlyList<string>? EventInvalidReleaseArguments { get; init; }
    public bool EventValidationInlineReturn { get; init; }
    public bool EventSinkOnly { get; init; }
    public string Source { get; init; } = "";
    public int Line { get; init; }
}
public sealed record VersionSupport { public string Product { get; init; } = ""; public IReadOnlyList<string> Versions { get; init; } = Array.Empty<string>(); }
public sealed record InvocationPlan
{
    public string Operation { get; init; } = "method";
    public ProjectionInvocationOperation OperationKind { get; init; } = ProjectionInvocationOperation.Method;
    public ProjectionInvocationApi Api { get; init; } = ProjectionInvocationApi.Factory;
    public string? Target { get; init; } = "this";
    public string DispatchName { get; init; } = "";
    public int? DispId { get; init; }
    public int ArgumentCount { get; init; }
    public IReadOnlyList<string> ArgumentOrder { get; init; } = Array.Empty<string>();
    public string? ResultType { get; init; }
    public ProjectedTypeReference? ResultTypeReference { get; init; }
    public ProjectionReturnConversion? SetValueConversion { get; init; }
    public ProjectionReturnConversion ReturnConversion { get; init; }
    public string? FactoryMethodSuffix { get; init; }
    public ProjectionArgumentPacking ArgumentPacking { get; init; }
    public ProjectionObjectArrayStyle ObjectArrayStyle { get; init; }
    public ProjectionInvocationCallKind CallKind { get; init; }
    public ProjectionReturnCastStyle ReturnCastStyle { get; init; }
    public ProjectionRawCallCast RawCallCast { get; init; }
    public ProjectionKnownReferenceFactoryStyle KnownReferenceFactoryStyle { get; init; }
    public ProjectionReturnValueStyle ReturnValueStyle { get; init; }
    public string? LocalReturnName { get; init; }
    public string? LocalReturnType { get; init; }
    public ProjectionInvokerCallStyle InvokerCallStyle { get; init; }
    public bool RequiresProxy { get; init; }
    public bool HasContractInvocation { get; init; }
    public IReadOnlyList<InvocationArgumentPlan> Arguments { get; init; } = Array.Empty<InvocationArgumentPlan>();
    public bool ReleaseArguments { get; init; }
}
public sealed record InvocationArgumentPlan
{
    public string Expression { get; init; } = "";
    public string? WriteBackExpression { get; init; }
    public string WriteBackType { get; init; } = "object";
    public bool ByRef { get; init; }
    public bool IsPropertyValue { get; init; }
    public string? InitializationExpression { get; init; }
    public string? WriteBackConversion { get; init; }
}
public sealed record ProjectionTrace
{
    public string LogicalId { get; init; } = "";
    public IReadOnlyList<ProjectionTraceStep> Steps { get; init; } = Array.Empty<ProjectionTraceStep>();
}
public sealed record ProjectionTraceStep { public string Stage { get; init; } = ""; public string Rule { get; init; } = ""; public string Result { get; init; } = ""; }
