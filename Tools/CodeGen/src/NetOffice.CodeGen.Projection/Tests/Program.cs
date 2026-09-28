using System.Text.Json;
using NetOffice.CodeGen.Data;
using NetOffice.CodeGen.Projection;

const string Sha = "0000000000000000000000000000000000000000000000000000000000000000";
var provenance = new Provenance { SourcePath = "fixture", SourceSha256 = Sha };
Provenance At(string location) => provenance with { Location = location + ":1" };
DataParameter P(string name, string type, bool optional = false, string? defaultValue = null, string refKind = "value") => new() { Name = name, Type = type, IsOptional = optional, HasDefaultValue = defaultValue is not null, DefaultValue = defaultValue, RefKind = refKind, Provenance = At("parameters/" + name) };

var libraries = new[] { "DAO", "Excel", "PowerPoint", "Outlook", "Access", "ADODB", "Office" }
    .Select((name, index) => new DataLibrary { LogicalId = "lib." + name, Name = name, SourceKey = "library." + name, Guid = $"00000000-0000-0000-0000-{index + 1:000000000000}", Version = "16.0", Major = "16", Minor = "0", Provenance = At(name + "/Project.xml") }).ToArray();
var projects = libraries.Select(library => new DataProject { LogicalId = "project." + library.Name, Name = library.Name, Namespace = "NetOffice." + library.Name + "Api", SourceKey = "project." + library.Name, Version = "1.0", FileVersion = "1.0", SourceCategories = new[] { "DispatchInterfaces", "Interfaces", "CoClasses", "Enums", "Modules", "Constants", "Records", "TypeDefs" }, LibraryIds = new[] { library.LogicalId }, Provenance = At(library.Name + "/Project.xml") }).ToArray();
DataType T(string product, string id, string name, string kind, string file, params string[] bases) => new() { LogicalId = id, LibraryId = "lib." + product, ProjectId = "project." + product, Namespace = "NetOffice." + product + "Api", SourceCategory = Path.GetFileNameWithoutExtension(file), Name = name, Kind = kind, SourceKey = id, BaseTypeIds = bases, Provenance = At(product + "/" + file) };
var types = new[]
{
    T("DAO", "dao._dao", "_DAO", "Interface", "DispatchInterfaces.xml"),
    T("DAO", "dao.database", "Database", "Interface", "DispatchInterfaces.xml", "dao._dao"),
    T("DAO", "dao.global", "Global", "Module", "Modules.xml"),
    T("DAO", "dao.language-constants", "LanguageConstants", "Constant", "Constants.xml"),
    T("PowerPoint", "powerpoint.presentation", "Presentation", "Interface", "DispatchInterfaces.xml"),
    T("Excel", "excel.range", "Range", "Interface", "DispatchInterfaces.xml"),
    T("PowerPoint", "powerpoint.presentations", "Presentations", "Interface", "DispatchInterfaces.xml"),
    T("PowerPoint", "powerpoint.application", "Application", "Interface", "DispatchInterfaces.xml"),
    T("PowerPoint", "powerpoint.eapplication", "EApplication", "Interface", "Interfaces.xml") with { IsEventInterface = true },
    T("Outlook", "outlook.application", "Application", "Interface", "DispatchInterfaces.xml"),
    T("Access", "access.application", "Application", "Interface", "DispatchInterfaces.xml"),
    T("ADODB", "adodb._recordset", "_Recordset", "Interface", "Interfaces.xml"),
    T("ADODB", "adodb.recordset", "Recordset", "CoClass", "CoClasses.xml", "adodb._recordset"),
    T("Office", "office.adjustments", "Adjustments", "Interface", "DispatchInterfaces.xml"),
    T("Office", "office.shape-type", "MsoAutoShapeType", "Enum", "Enums.xml"),
    T("Office", "office.constants", "MsoConstants", "Constant", "Constants.xml"),
    T("Office", "office.point", "Point", "Record", "Records.xml"),
    T("Office", "office.rgb", "MsoRGBType", "Alias", "TypeDefs.xml")
};

var members = new List<DataMember>();
var invocation = new List<InvocationEvidence>();
void M(string id, string owner, string name, string kind, string? result, int dispId, string? accessor = null, string? group = null, params DataParameter[] parameters)
{
    var evidenceId = "invoke." + id;
    members.Add(new DataMember { LogicalId = id, TypeId = owner, Name = name, Kind = kind, SourceKey = id, DispId = dispId, ReturnType = result, ParameterTypes = parameters.Select(x => x.Type).ToArray(), Parameters = parameters, AccessorKind = accessor, AccessorGroupId = group, InvocationEvidenceId = evidenceId, Provenance = At(owner + "/" + name) });
    invocation.Add(new InvocationEvidence { LogicalId = evidenceId, MemberId = id, Operation = kind, DispatchName = name, DispId = dispId, ArgumentCount = parameters.Length, ResultType = result, RequiresProxy = result is not null && !new[] { "void", "string", "Int32", "bool" }.Contains(result), Provenance = At(owner + "/" + name) });
}
M("dao.database.name", "dao.database", "Name", "property", "string", 1, "get", "group.dao.name");
M("dao.database.connect.get", "dao.database", "Connect", "property", "string", 2, "get", "group.dao.connect");
M("dao.database.connect.putref", "dao.database", "Connect", "property", "string", 2, "putref", "group.dao.connect");
M("dao.database.open", "dao.database", "OpenRecordset", "method", "Recordset", 3, null, null, P("name", "string"), P("type", "object", true), P("options", "object", true));
M("dao.global.opendatabase", "dao.global", "OpenDatabase", "method", "Database", 4, null, null, P("name", "string"));
M("excel.range.calculate", "excel.range", "Calculate", "method", "object", 10, null, null, P("rowMajorOrder", "object"), P("iteration", "object", true, "-1"), P("maxChange", "object", true, "0.001"));
M("powerpoint.presentations.count", "powerpoint.presentations", "Count", "property", "Int32", 11);
M("powerpoint.presentations.item", "powerpoint.presentations", "Item", "property", "Presentation", 0, null, null, P("index", "object"));
M("powerpoint.eapplication.close", "powerpoint.eapplication", "PresentationClose", "event", "void", 2004, null, null, P("presentation", "Presentation"));
M("powerpoint.presentations.enum", "powerpoint.presentations", "GetEnumerator", "method", "IEnumerator", -4);
M("powerpoint.application.close", "powerpoint.application", "PresentationClose", "event", "void", 2004, null, null, P("presentation", "Presentation"));
M("outlook.application.send", "outlook.application", "ItemSend", "event", "void", 61442, null, null, P("item", "object"), P("cancel", "object", false, null, "ref"));
M("access.application.send", "access.application", "ItemSend", "event", "void", 61442, null, null, P("item", "object"), P("cancel", "object", false, null, "ref"));
M("adodb.recordset.open.1", "adodb.recordset", "Open", "method", "void", 20, null, null, P("source", "object"));
M("adodb.recordset.open.2", "adodb.recordset", "Open", "method", "void", 21, null, null, P("source", "object"));
M("office.adjustments.count", "office.adjustments", "Count", "property", "Int32", 1);
M("office.adjustments.item", "office.adjustments", "Item", "property", "Single", 0, "get", "group.office.item", P("index", "Int32", refKind: "ref"));

var groups = new[]
{
    new AccessorGroup { LogicalId = "group.dao.name", TypeId = "dao.database", Name = "Name", MemberIds = new[] { "dao.database.name" }, Provenance = At("DAO/DispatchInterfaces.xml") },
    new AccessorGroup { LogicalId = "group.dao.connect", TypeId = "dao.database", Name = "Connect", MemberIds = new[] { "dao.database.connect.get", "dao.database.connect.putref" }, Provenance = At("DAO/DispatchInterfaces.xml") },
    new AccessorGroup { LogicalId = "group.office.item", TypeId = "office.adjustments", Name = "Item", MemberIds = new[] { "office.adjustments.item" }, Provenance = At("Office/DispatchInterfaces.xml") }
};
var observations = new[]
{
    new SupportObservation { LogicalId = "support.dao.database", TargetId = "dao.database", Product = "DAO", Versions = new[] { "3.6" }, LibraryIds = new[] { "lib.DAO" }, Provenance = At("DAO/DispatchInterfaces.xml") },
    new SupportObservation { LogicalId = "support.dao.open", TargetId = "dao.database.open", Product = "DAO", Versions = new[] { "12.0" }, LibraryIds = new[] { "lib.DAO" }, Provenance = At("DAO/DispatchInterfaces.xml") },
    new SupportObservation { LogicalId = "support.office.enum", TargetId = "office.shape.value", Product = "Office", Versions = new[] { "16" }, LibraryIds = new[] { "lib.Office" }, Provenance = At("Office/Enums.xml") }
};
var values = new[]
{
    new DataValue { LogicalId = "office.shape.value", TypeId = "office.shape-type", Name = "msoShapeMixed", Kind = "Enum", Value = "-2", ValueType = "Int32", Provenance = At("Office/Enums.xml") },
    new DataValue { LogicalId = "office.shape.value.duplicate-observation", TypeId = "office.shape-type", Name = "msoShapeMixed", Kind = "Enum", Value = "0", ValueType = "Int32", Provenance = At("Office/Enums-duplicate.xml") },
    new DataValue { LogicalId = "office.constant.value", TypeId = "office.constants", Name = "msoTrue", Kind = "Constant", Value = "-1", ValueType = "Int32", Provenance = At("Office/Constants.xml") },
    new DataValue { LogicalId = "dao.language.value", TypeId = "dao.language-constants", Name = "dbLangGeneral", Kind = "Constant", Value = ";LANGID=0x0409;CP=1252;COUNTRY=0", ValueType = "string", Provenance = At("DAO/Constants.xml") },
    new DataValue { LogicalId = "office.point.x", TypeId = "office.point", Name = "X", Kind = "field", Value = "0", ValueType = "Int32", Provenance = At("Office/Records.xml") }
};
var graph = new DataGraph
{
    Source = new DataSource { Kind = "fixture", Path = "representative-products", Sha256 = Sha },
    Projects = projects,
    Libraries = libraries,
    Types = types,
    Members = members,
    Values = values,
    AccessorGroups = groups,
    SupportObservations = observations,
    InvocationEvidence = invocation
};
graph = graph with { Digest = CanonicalJson.ComputeDigest(graph) };

var databaseMembers = new[]
{
    CMember("Name", "property", "string", null, "public string Name {"),
    CMember("Connect", "property", "string", null, "public string Connect {") with { InvocationText = new[] { "Factory.ExecuteValuePropertySet(this, \"Connect\", value);", "return Factory.ExecuteStringPropertyGet(this, \"Connect\");" } },
    CMember("ContractOnly", "property", "string", null, "public string ContractOnly {") with { InvocationText = new[] { "Factory.ExecuteValuePropertySet(this, \"ContractOnly\", value);", "return Factory.ExecuteStringPropertyGet(this, \"ContractOnly\");" } },
    CMember("GetActiveInstance", "method", "Database", "(bool throwExceptionIfNotFound = false)", "public static Database GetActiveInstance(bool throwExceptionIfNotFound = false) {") with { Modifiers = new[] { "static" } },
    CMember("OpenRecordset", "method", "Recordset", "(string name, object type, object options)", "public Recordset OpenRecordset(string name, object type, object options)") with { Attributes = new[] { "Redirect(\"get_OpenRecordset\")" } },
    CMember("PackedCall", "method", "object", "(object left, object right)", "public object PackedCall(object left, object right) {") with { InvocationText = new[] { "return Factory.ExecuteObjectMethodGet(this, \"PackedCall\", new object[]{ left, right, 1 });" } },
    CMember("KnownCall", "method", "Recordset", "(object left)", "public Recordset KnownCall(object left) {") with { InvocationText = new[] { "return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, \"KnownCall\", NetOffice.DAOApi.Recordset.LateBindingApiWrapperType, left, 1);" } },
};
var contract = new WrapperContract
{
    SchemaVersion = "1.0",
    ContractKind = "NetOffice.WrapperContract",
    Generator = new ContractGenerator { Name = "fixture", Version = "1" },
    Source = new ContractSource { Root = "Source", Api = "DAO", SourceRoots = new[] { "Source/DAO" } },
    Types = new[]
    {
        new ContractType { LogicalId = "NetOffice.DAOApi.Database", Namespace = "NetOffice.DAOApi", Name = "Database", Kind = "class", BaseType = "_DAO", Attributes = new[] { "SupportByVersion(\"DAO\", 3.6,12.0)", "EntityType(EntityType.IsDispatchInterface)" }, Source = "DispatchInterfaces/Database.cs", Line = 14, Signature = "public class Database : _DAO", Members = databaseMembers },
        new ContractType
        {
            LogicalId = "DAOApi.Utils.ProjectInfo",
            Namespace = "DAOApi.Utils",
            Name = "ProjectInfo",
            Kind = "class",
            BaseType = "IFactoryInfo",
            Source = "Utils/ProjectInfo.cs",
            Line = 12,
            Signature = "public class ProjectInfo : IFactoryInfo",
            Members = new[] { new ContractMember { Name = "ProjectInfo", Kind = "constructor", Signature = "public ProjectInfo() {", Source = "Utils/ProjectInfo.cs", Line = 28 } }
        }
    }
};
var policy = ProjectionPolicy.Default;
var result = ProjectionEngine.Project(graph, contract, policy);
var resultAgain = ProjectionEngine.Project(graph, contract, policy);
var resultWithoutTraces = ProjectionEngine.Project(graph, contract, policy, new ProjectionOptions { IncludeTraces = false });


Check(result.ToJson() == resultAgain.ToJson(), "projection must be deterministic");
Check(result.Files.SelectMany(file => file.Types).Count(type => !string.IsNullOrEmpty(type.DataLogicalId)) == result.Coverage.DataTypes, "all Data types must be projected");
Check(result.Coverage.DataMembers == result.Coverage.ProjectedDataMembers, "all Data members must be projected");
Check(result.Coverage.DataValues == result.Coverage.ProjectedValues, "all Data values must be projected");
Check(resultWithoutTraces.Explain.Count == 0 && resultWithoutTraces.Files.Count == result.Files.Count,
    "trace-free projection skips explain evidence without changing projected files");

var projectedProjectInfo = result.Files.SelectMany(file => file.Types).Single(type => type.ProjectInfo is not null);
Check(projectedProjectInfo.EntityKind == ProjectionEntityKind.Utility && projectedProjectInfo.FileCategory == ProjectionFileCategory.Utils, "ProjectInfo is a typed generated utility");
Check(projectedProjectInfo.ProjectInfo!.AssemblyNamespace == "NetOffice.DAOApi", "ProjectInfo assembly namespace is derived from Data v2");
Check(projectedProjectInfo.ProjectInfo.ComponentGuids.SequenceEqual(new[] { "00000000-0000-0000-0000-000000000001" }), "ProjectInfo component GUIDs are derived from project libraries");
Check(projectedProjectInfo.Source == "Utils/ProjectInfo.cs", "ProjectInfo preserves its contract partition");

var projectedConnectGetters = result.Files.SelectMany(file => file.Types).Single(type => type.DataLogicalId == "dao.database").Members
    .Where(member => member.CSharpName == "Connect").SelectMany(member => member.AccessorInvocations)
    .Where(invocationPlan => invocationPlan.OperationKind == ProjectionInvocationOperation.PropertyGet).ToArray();
Check(projectedConnectGetters.Length != 0 && projectedConnectGetters.All(invocationPlan => invocationPlan.Arguments.All(argument => argument.Expression != "value")), "synthesized getter never inherits the setter value argument");
Check(projectedConnectGetters.All(invocationPlan => invocationPlan.FactoryMethodSuffix == "String"), "exact scalar factory suffix is typed in the invocation plan");
var packedCall = result.Files.SelectMany(file => file.Types).Single(type => type.DataLogicalId == "dao.database").Members.Single(member => member.CSharpName == "PackedCall");
Check(packedCall.Invocation.ArgumentPacking == ProjectionArgumentPacking.ObjectArray
      && packedCall.Invocation.ObjectArrayStyle == ProjectionObjectArrayStyle.Compact
      && packedCall.Invocation.CallKind == ProjectionInvocationCallKind.MethodGet
      && packedCall.Invocation.HasContractInvocation
      && packedCall.Invocation.Arguments.Select(static argument => argument.Expression).SequenceEqual(new[] { "left", "right", "1" }, StringComparer.Ordinal),
    "exact object-array dispatch packing unwraps its payload once and types the call style: " + string.Join(" | ", packedCall.Invocation.Arguments.Select(static argument => argument.Expression)));
var knownCall = result.Files.SelectMany(file => file.Types).Single(type => type.DataLogicalId == "dao.database").Members.Single(member => member.CSharpName == "KnownCall");
Check(knownCall.Invocation.Arguments.Select(static argument => argument.Expression).SequenceEqual(new[] { "left", "1" }, StringComparer.Ordinal),
    "exact known-reference payload excludes the helper wrapper-type operand while preserving fixed COM arguments");
var sinkOnlyContract = new WrapperContract
{
    SchemaVersion = "1.0",
    ContractKind = "NetOffice.WrapperContract",
    Generator = new ContractGenerator { Name = "fixture", Version = "1" },
    Source = new ContractSource { Root = "Source", Api = "PowerPoint", SourceRoots = new[] { "Source/PowerPoint" } },
    Types = new[]
    {
        new ContractType
        {
            LogicalId = "NetOffice.PowerPointApi.Events.EApplication",
            Namespace = "NetOffice.PowerPointApi.Events",
            Name = "EApplication",
            Kind = "interface",
            Source = "Events/EApplication.cs",
            Line = 10,
            Signature = "public interface EApplication",
            Members = Array.Empty<ContractMember>()
        },
        new ContractType
        {
            LogicalId = "NetOffice.PowerPointApi.Events.EApplicationFixture_SinkHelper",
            Namespace = "NetOffice.PowerPointApi.Events",
            Name = "EApplicationFixture_SinkHelper",
            Kind = "class",
            Source = "Events/EApplication.cs",
            Line = 30,
            Signature = "internal class EApplicationFixture_SinkHelper",
            Members = new[]
            {
                CMember("PresentationClose", "method", "void", "(NetOffice.PowerPointApi.Presentation presentation)", "public void PresentationClose(NetOffice.PowerPointApi.Presentation presentation)") with
                {
                    InvocationText = new[] { "Invoker.ReleaseParamsArray(presentation);" },
                    Source = "Events/EApplication.cs",
                    Line = 50
                },
                CMember("QuotedValidation", "method", "void", "(object presentation)", "public void QuotedValidation(object presentation)") with
                {
                    InvocationText = new[] { "Validate(\"Invoker.ReleaseParamsArray(presentation);\");" },
                    Source = "Events/EApplication.cs",
                    Line = 60
                }
            }
        }
    }
};
var sinkOnlyMembers = ProjectionEngine.Project(graph, sinkOnlyContract, policy).Files.SelectMany(file => file.Types)
    .SelectMany(type => type.Members).Where(member => member.CSharpName == "PresentationClose"
                                                      && member.EmissionOwnerLogicalId.StartsWith("NetOffice.PowerPointApi.Events.EApplication", StringComparison.Ordinal)).ToArray();
Check(sinkOnlyMembers.Count(member => member.EventSinkOnly && member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit) == 1
      && sinkOnlyMembers.All(member => member.EventSinkOnly || member.EmissionDisposition == ProjectionMemberEmissionDisposition.SuppressedByContract)
      && sinkOnlyMembers.Single(member => member.EventSinkOnly).EventInvalidReleaseArguments?.SequenceEqual(new[] { "presentation" }, StringComparer.Ordinal) == true,
    "sink-helper-only contract members stay out of the event interface and carry an explicit emission flag: "
    + string.Join(" | ", sinkOnlyMembers.Select(member => $"{member.Source}:{member.EmissionOwnerLogicalId}:{member.EmissionDisposition}:{member.EventSinkOnly}")));
var quotedValidation = ProjectionEngine.Project(graph, sinkOnlyContract, policy).Files.SelectMany(file => file.Types)
    .SelectMany(type => type.Members).Single(member => member.CSharpName == "QuotedValidation");
Check(quotedValidation.EventValidationKey == "Invoker.ReleaseParamsArray(presentation);"
      && quotedValidation.EventInvalidReleaseArguments is null,
    "release-call text inside a validation string is not projected as an executable release");
var eventBindingContract = new WrapperContract
{
    SchemaVersion = "1.0",
    ContractKind = "NetOffice.WrapperContract",
    Generator = new ContractGenerator { Name = "fixture", Version = "1" },
    Source = new ContractSource { Root = "Source", Api = "ADODB", SourceRoots = new[] { "Source/ADODB" } },
    Types = new[]
    {
        new ContractType
        {
            LogicalId = "NetOffice.ADODBApi.Recordset",
            Namespace = "NetOffice.ADODBApi",
            Name = "Recordset",
            Kind = "class",
            Attributes = new[] { "EntityType(EntityType.IsCoClass)", "EventSink(typeof(Events.RecordsetEvents_SinkHelper))" },
            Source = "Classes/Recordset.cs",
            Line = 1,
            Signature = "public class Recordset : _Recordset",
            Members = new[]
            {
                new ContractMember
                {
                    Name = "WillChangeRecord",
                    Kind = "event",
                    ReturnType = "RecordsetEvents_WillChangeRecordEventHandler",
                    Signature = "public event RecordsetEvents_WillChangeRecordEventHandler WillChangeRecord {",
                    InvocationText = new[] { "_willChangeRecordEvent += value;", "_willChangeRecordEvent -= value;" },
                    Source = "Classes/Recordset.cs",
                    Line = 40
                }
            }
        }
    }
};
var eventBinding = ProjectionEngine.Project(graph, eventBindingContract, policy).Files.SelectMany(file => file.Types).Single(type => type.DataLogicalId == "adodb.recordset").EventBindings.Single();
Check(eventBinding.SinkHelperType == "Events.RecordsetEvents_SinkHelper", "coclass event binding preserves qualified sink helper type");
Check(eventBinding.FieldName == "_recordsetEvents_SinkHelper", "coclass event binding supplies deterministic field name");
var customEvent = ProjectionEngine.Project(graph, eventBindingContract, policy).Files.SelectMany(file => file.Types).Single(type => type.DataLogicalId == "adodb.recordset").Members.Single(member => member.CSharpName == "WillChangeRecord");
Check(customEvent.UseCustomEventAccessors && customEvent.EventBackingField == "_willChangeRecordEvent", "exact event accessor shape and backing field are typed");
var excelContract = new WrapperContract
{
    SchemaVersion = "1.0",
    ContractKind = "NetOffice.WrapperContract",
    Generator = new ContractGenerator { Name = "fixture", Version = "1" },
    Source = new ContractSource { Root = "Source", Api = "Excel", SourceRoots = new[] { "Source/Excel" } },
    Types = new[]
    {
        new ContractType { LogicalId = "NetOffice.ExcelApi.Native.Range", Namespace = "NetOffice.ExcelApi.Native", Name = "Range", Kind = "interface", Source = "Native/Range.cs", Line = 1, Signature = "public interface Range" },
        new ContractType
        {
            LogicalId = "NetOffice.ExcelApi.Range",
            Namespace = "NetOffice.ExcelApi",
            Name = "Range",
            Kind = "class",
            BaseType = "COMObject",
            Attributes = new[] { "FragmentOne" },
            Source = "DispatchInterfaces/Range.First.cs",
            Line = 10,
            Signature = "public partial class Range : COMObject",
            Members = new[] { new ContractMember { Name = "Calculate", Kind = "method", ReturnType = "string", Parameters = "(object rowMajorOrder, object iteration, object maxChange)", Signature = "public string Calculate(object rowMajorOrder, object iteration, object maxChange)", Source = "DispatchInterfaces/Range.First.cs", Line = 20 } }
        },
        new ContractType
        {
            LogicalId = "NetOffice.ExcelApi.DispatchInterfaces.Range",
            Namespace = "NetOffice.ExcelApi.DispatchInterfaces",
            Name = "Range",
            Kind = "class",
            BaseType = "COMObject",
            Attributes = new[] { "FragmentTwo" },
            Source = "DispatchInterfaces/Range.Second.cs",
            Line = 30,
            Signature = "public class Range : COMObject",
            Members = new[] { new ContractMember { Name = "Calculate", Kind = "method", ReturnType = "bool", Parameters = "(object rowMajorOrder, object iteration)", Signature = "public bool Calculate(object rowMajorOrder, object iteration)", Source = "DispatchInterfaces/Range.Second.cs", Line = 40 } }
        }
    }
};
var excelResult = ProjectionEngine.Project(graph, excelContract, policy);
var rangeParts = excelResult.Files.SelectMany(file => file.Types).Where(type => type.DataLogicalId == "excel.range").ToArray();
var mergedRange = rangeParts.Single(type => type.IsPrimaryPart);
var rangeMembers = rangeParts.SelectMany(type => type.Members).ToArray();
Check(mergedRange.LogicalId == "NetOffice.ExcelApi.Range", "canonical contract identity wins over native and partial fragments");
Check(rangeParts.SelectMany(type => type.Attributes).Contains("FragmentOne") && rangeParts.SelectMany(type => type.Attributes).Contains("FragmentTwo"), "partial contract metadata is partitioned");
Check(rangeMembers.Single(member => member.DataLogicalId == "excel.range.calculate" && member.ParameterList.Count == 3).ReturnType == "string", "first partial members remain matchable");
Check(rangeMembers.Single(member => member.DataLogicalId == "excel.range.calculate" && member.ParameterList.Count == 2).ReturnType == "bool", "second partial members remain matchable");
Check(rangeParts.Single(type => type.Namespace == "NetOffice.ExcelApi.DispatchInterfaces").ContractPartLogicalId == "NetOffice.ExcelApi.DispatchInterfaces.Range", "partial contract owner identity is explicit");

var overridePolicy = ProjectionPolicy.WithDigest(policy with
{
    Overrides = new[] { new ProjectionOverride { Id = "range-calculate-name", TypeLogicalId = "excel.range", MemberLogicalId = "excel.range.calculate", Property = "csharpName", Value = "CalculateProjected", ExpectedMatches = 3 } }
});
var overridden = ProjectionEngine.Project(graph, contract, overridePolicy);
var overriddenRange = overridden.Files.SelectMany(file => file.Types).Single(type => type.DataLogicalId == "excel.range");
Check(overriddenRange.Members.Where(member => member.DataLogicalId == "excel.range.calculate").All(member => member.CSharpName == "CalculateProjected"), "typed member override");
Check(overridden.Explain.Where(trace => trace.LogicalId.StartsWith("excel.range/excel.range.calculate", StringComparison.Ordinal)).All(trace => trace.Steps.Any(step => step.Stage == "policy" && step.Rule == "typed-override:range-calculate-name")), "override explain trace");

using (var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "fixtures", "representative-expectations.json"))))
{
    Check(fixture.RootElement.GetProperty("schemaVersion").GetString() == "projection-fixture/v1", "representative fixture schema");
    foreach (var expected in fixture.RootElement.GetProperty("entities").EnumerateArray())
    {
        var scenario = expected.GetProperty("scenario").GetString()!;
        var entity = Type(expected.GetProperty("logicalId").GetString()!);
        Check(entity.Namespace == expected.GetProperty("namespace").GetString(), scenario + " namespace");
        Check(entity.FileCategory.ToString() == expected.GetProperty("category").GetString(), scenario + " category");
        var capability = expected.GetProperty("capability");
        if (capability.ValueKind == JsonValueKind.String) Check(entity.Capabilities.Contains(capability.GetString()!), scenario + " capability");
    }
}

var database = Type("dao.database");
Check(database.Namespace == "NetOffice.DAOApi", "DAO namespace");
Check(database.Kind == "class" && database.EntityKind == ProjectionEntityKind.DispatchInterface, "DAO Database entity kind");
Check(database.BaseType == "_DAO", "DAO Database base type");
Check(database.RuntimeRequirements.Contains("netoffice-core") && database.RuntimeRequirements.Contains("com-proxy") && database.Constructors.Count >= 7, "DAO Database runtime contract");
Check(database.SupportVersions.Single(x => x.Product == "DAO").Versions.SequenceEqual(new[] { "3.6", "12.0" }), "DAO support observations and contract support");
Check(database.Members.Any(x => x.Name == "Name" && x.Invocation.OperationKind == ProjectionInvocationOperation.PropertyGet), "DAO property get");
var connectPutRef = database.Members.Single(x => x.Name == "Connect" && x.Invocation.OperationKind == ProjectionInvocationOperation.PropertyPutRef);
Check(connectPutRef.Invocation.ArgumentOrder.SequenceEqual(new[] { "value" }) && connectPutRef.Invocation.ArgumentCount == 1 && connectPutRef.Invocation.Arguments.Single().IsPropertyValue && connectPutRef.Invocation.ReturnConversion == ProjectionReturnConversion.None && connectPutRef.Invocation.SetValueConversion == ProjectionReturnConversion.Value, "DAO putref invocation");
var redirectedOpen = database.Members.Single(member => member.Name == "OpenRecordset" && member.ParameterList.Count == 3);
Check(redirectedOpen.Invocation.OperationKind == ProjectionInvocationOperation.LocalForward && redirectedOpen.Invocation.Api == ProjectionInvocationApi.Local && redirectedOpen.Invocation.Target is null && redirectedOpen.Invocation.DispatchName == "get_OpenRecordset" && redirectedOpen.Invocation.ArgumentOrder.SequenceEqual(new[] { "name", "type", "options" }), "redirect attributes project a typed local forward");
var contractOnlyProperty = database.Members.Single(member => member.Name == "ContractOnly");
Check(contractOnlyProperty.AccessorInvocations.Select(static invocationPlan => invocationPlan.OperationKind).SequenceEqual(new[] { ProjectionInvocationOperation.PropertyGet, ProjectionInvocationOperation.PropertySet }), "contract-only property has distinct getter and setter plans");
var activeInstance = database.Members.Single(member => member.Name == "GetActiveInstance");
Check(activeInstance.Invocation.OperationKind == ProjectionInvocationOperation.ActiveInstance && activeInstance.Invocation.Api == ProjectionInvocationApi.Runtime && activeInstance.Invocation.Target is null, "static active instance helpers use the typed runtime plan");
Check(Type("dao._dao").EmissionDisposition == ProjectionEmissionDisposition.SuppressedByContract, "closed wrapper contract suppresses unmatched Data types without dropping their projection trace");

var range = Type("excel.range");
var calculate = range.Members.Where(x => x.Name == "Calculate").OrderByDescending(x => x.ParameterList.Count).ToArray();
Check(calculate.Select(x => x.ParameterList.Count).SequenceEqual(new[] { 3, 2, 1 }), "Excel optional overload family");
Check(calculate[0].ParameterList[1].HasDefaultValue && calculate[0].ParameterList[1].DefaultValue == "-1", "Excel optional default preserved");
Check(calculate.All(member => member.ParameterList.All(parameter => !parameter.EmitDefaultValue)), "expanded overloads do not emit C# defaults");
Check(calculate[0].Invocation.ReturnConversion == ProjectionReturnConversion.UntypedReference, "untyped object returns use the non-generic reference invocation");
Check(calculate[0].Invocation.ArgumentOrder.SequenceEqual(new[] { "rowMajorOrder", "iteration", "maxChange" }), "Excel invocation argument order");

var presentations = Type("powerpoint.presentations");
Check(presentations.Capabilities.Contains("collection") && presentations.Capabilities.Contains("enumerator") && presentations.Capabilities.Contains("indexer"), "PowerPoint collection capabilities");
Check(presentations.Members.Single(x => x.Name == "Item").Invocation.ReturnConversion == ProjectionReturnConversion.KnownReference, "PowerPoint known proxy return conversion");
Check(Type("powerpoint.application").Capabilities.Contains("event"), "PowerPoint event source capability");
Check(Type("powerpoint.eapplication").EntityKind == ProjectionEntityKind.EventInterface && Type("powerpoint.eapplication").Namespace == "NetOffice.PowerPointApi.Events" && Type("powerpoint.eapplication").EventSink?.Name == "EApplication_SinkHelper", "PowerPoint event interface and sink projection");
Check(Type("outlook.application").Namespace == "NetOffice.OutlookApi" && Type("access.application").Namespace == "NetOffice.AccessApi", "same-named event owners stay product-scoped");
Check(Type("outlook.application").Members.Single().DataLogicalId != Type("access.application").Members.Single().DataLogicalId, "event conflict identities stay distinct");

var recordset = Type("adodb.recordset");
Check(recordset.Constructors.Any(x => x.Kind == "activation" && x.BaseCall == "base(\"ADODB.Recordset\")"), "ADODB CoClass activation constructor");
Check(recordset.EffectiveBaseTypes.Any(x => x.EndsWith("._Recordset", StringComparison.Ordinal)), "ADODB inheritance closure");
Check(recordset.Members.Count(x => x.DuplicateOf is not null) == 1, "ADODB duplicate signatures are explicit");
var adjustments = Type("office.adjustments");
Check(adjustments.Capabilities.Contains("indexer") && adjustments.Members.Any(x => x.Kind == "indexer" && x.ParameterList.All(parameter => parameter.RefKind == "value")), "Office adjustment indexer normalizes forbidden ref parameters");
Check(Type("dao.language-constants").Members.Single().ConstantExpression == "\";LANGID=0x0409;CP=1252;COUNTRY=0\"", "string constant literal escaping");
Check(Type("office.shape-type").Members.Single(x => x.DataLogicalId == "office.shape.value").Kind == "enum-value", "enum values projected");
Check(Type("office.shape-type").Members.Count(member => member.CSharpName == "msoShapeMixed" && member.EmissionDisposition == ProjectionMemberEmissionDisposition.Emit) == 1, "duplicate enum observations emit one exact declaration");
Check(Type("office.constants").FileCategory == ProjectionFileCategory.Constants && Type("office.constants").Members.Single().ConstantValue == "-1", "constants projected in current category");
Check(Type("office.rgb").Attributes.Contains("EntityType(EntityType.IsStruct)"), "typedefs use a valid runtime entity attribute");
Check(Type("dao.global").Modifiers.Contains("static"), "modules project as static classes");
Check(result.Files.All(x => x.Path.StartsWith("Generated/", StringComparison.Ordinal)), "generated partition root");
Console.WriteLine($"Projection fixtures passed: {result.Coverage.ProjectedTypes} types, {result.Coverage.ProjectedDataMembers} members, {result.Coverage.ProjectedValues} values.");

if (args.Length == 2)
{
    var fullGraph = CanonicalJson.Read(args[0]);
    var fullContract = WrapperContract.Read(args[1]);
    var full = ProjectionEngine.Project(fullGraph, fullContract, ProjectionPolicy.Default);
    Check(full.Coverage.ProjectedTypes >= fullGraph.Types.Count, "full graph type coverage");
    Check(full.Coverage.ProjectedDataMembers == fullGraph.Members.Count, "full graph member coverage");
    Check(full.Coverage.ProjectedValues == fullGraph.Values.Count, "full graph value coverage");
    Check(full.Files.Select(file => file.Path).Distinct(StringComparer.Ordinal).Count() == full.Files.Count, "full graph unique partitions");
    var fullMembers = full.Files.SelectMany(file => file.Types).SelectMany(type => type.Members).ToArray();
    Check(fullMembers.Where(member => member.Kind == "indexer").SelectMany(member => member.ParameterList).All(parameter => parameter.RefKind == "value"), "full graph indexers contain no forbidden ref/out parameters");
    Check(fullMembers.SelectMany(member => member.ParameterList).Where(parameter => !parameter.IsContractOverlay).All(parameter => !parameter.EmitDefaultValue), "full graph Data overloads emit no invalid optional defaults");
    Check(fullMembers.SelectMany(member => member.ParameterList).Select(parameter => parameter.TypeReference).Where(reference => reference?.TargetTypeId is not null).All(reference => !string.IsNullOrWhiteSpace(reference!.QualifiedName) && reference.QualifiedName!.Contains('.', StringComparison.Ordinal)), "full graph resolved references have qualified C# names");
    Check(full.Files.SelectMany(file => file.Types).Where(type => type.EntityKind == ProjectionEntityKind.EventInterface).All(type => type.EventSink is not null), "full graph event interfaces include sink plans");
    Check(fullMembers.SelectMany(member => member.ParameterList).Select(parameter => parameter.SinkArgument).Where(argument => argument is not null)
        .All(argument => argument!.Conversion != ProjectionEventConversion.Scalar || argument.ConversionExpression?.Contains("{0}", StringComparison.Ordinal) == true), "full graph scalar event arguments have conversion templates");
    Check(fullMembers.SelectMany(member => member.ParameterList).Select(parameter => parameter.SinkArgument).Where(argument => argument?.WriteBackExpression is not null)
        .All(argument => argument!.WriteBackExpression!.Contains("{0}", StringComparison.Ordinal)), "full graph event writeback expressions are templates");
    Check(fullMembers.SelectMany(member => member.AccessorInvocations.Append(member.Invocation))
        .All(invocation => invocation.Arguments.All(static argument => !argument.ByRef) || invocation.Api == ProjectionInvocationApi.Invoker), "full graph by-ref invocation plans use Invoker");
    if (string.Equals(fullContract.Source.Api, "Office", StringComparison.Ordinal))
    {
        Check(full.Files.SelectMany(file => file.Types).Single(type => type.Product == "Office" && type.Name == "IRibbonControl").EmissionDisposition == ProjectionEmissionDisposition.Companion, "native contract ownership suppresses duplicate generated wrappers");
        Check(full.Files.SelectMany(file => file.Types).Single(type => type.Product == "Office" && type.Name == "IRibbonUI").EmissionDisposition == ProjectionEmissionDisposition.Companion, "native caller contract ownership suppresses duplicate generated wrappers");
        var accessible = full.Files.SelectMany(file => file.Types).Single(type => type.Product == "Office" && type.Name == "IAccessible");
        Check(accessible.AuxiliaryTypes.Any(type => type.Name == "IAccessible_"), "contract-only syntax-bypass base type is projected");
        Check(accessible.Members.Any(member => member.Name == "accName" && member.Kind == "method" && member.EmissionOwnerLogicalId.EndsWith(".IAccessible_", StringComparison.Ordinal)), "syntax-bypass overload ownership is preserved");
        var customTaskPane = full.Files.SelectMany(file => file.Types).Single(type => type.Product == "Office" && type.Name == "CustomTaskPane");
        Check(customTaskPane.AuxiliaryTypes.Count(type => type.Kind == "delegate" && type.DelegateParameters.Count == 1) == 2, "contract-only event delegates are projected as typed auxiliary declarations");
    }
    var fullDatabase = full.Files.SelectMany(file => file.Types).Single(type => type.Product == "DAO" && type.Name == "Database");
    Check(fullDatabase.Namespace == "NetOffice.DAOApi" && fullDatabase.BaseType?.EndsWith("._DAO", StringComparison.Ordinal) == true, "full graph DAO Database parity");
    Check(fullDatabase.EntityKind == ProjectionEntityKind.DispatchInterface && fullDatabase.RuntimeRequirements.Contains("com-proxy") && fullDatabase.Constructors.Count == 8, "full graph DAO Database wrapper runtime");
    Check(fullDatabase.SupportVersions.Single(item => item.Product == "DAO").Versions.SequenceEqual(new[] { "3.6", "12.0" }), "full graph DAO Database support facts");
    Check(fullDatabase.Members.Any(member => member.Name == "Name" && member.Invocation.OperationKind == ProjectionInvocationOperation.PropertyGet), "full graph DAO Database property");
    Check(fullDatabase.Members.Any(member => member.Name == "OpenRecordset" && member.Invocation.OperationKind == ProjectionInvocationOperation.Method), "full graph DAO Database method");
    var fullDatabaseData = fullGraph.Types.Single(type => type.LogicalId == fullDatabase.DataLogicalId);
    var fullDatabaseMemberIds = fullGraph.Members.Where(member => member.TypeId == fullDatabaseData.LogicalId).Select(member => member.LogicalId).ToHashSet(StringComparer.Ordinal);
    Check(fullDatabaseMemberIds.All(id => fullDatabase.Members.Any(member => member.DataLogicalId == id)), "full graph DAO Database retains every Data-only member");
    Console.WriteLine($"Full graph passed: {full.Coverage.ProjectedTypes} types, {full.Coverage.ProjectedDataMembers} members, {full.Coverage.ProjectedValues} values.");
}

ProjectedType Type(string id) => result.Files.SelectMany(x => x.Types).Single(x => x.DataLogicalId == id);
static ContractMember CMember(string name, string kind, string? returnType, string? parameters, string signature) => new() { Name = name, Kind = kind, ReturnType = returnType, Parameters = parameters, Signature = signature, Source = "DispatchInterfaces/Database.cs", Line = 100 };
static void Check(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
