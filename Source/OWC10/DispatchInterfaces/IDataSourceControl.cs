using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// IDataSourceControl
    /// </summary>
    [SyntaxBypass]
    public interface IDataSourceControl_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="dataMember">optional object dataMember</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api.Enums.ProviderType get_ProviderType(object dataMember);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_ProviderType
        /// </summary>
        /// <param name="dataMember">optional object dataMember</param>
        [SupportByVersion("OWC10", 1), Redirect("get_ProviderType")]
        NetOffice.OWC10Api.Enums.ProviderType ProviderType(object dataMember);

        #endregion
    }

    /// <summary>
    /// DispatchInterface IDataSourceControl 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("F5B39AAF-1480-11D3-8549-00C04FAC67D7")]
    [CoClassSource(typeof(NetOffice.OWC10Api.DataSourceControl))]
    public interface IDataSourceControl : IDataSourceControl_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string ConnectionString { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string CurrentDirectory { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool UseRemoteProvider { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.ADODBApi.Connection Connection { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool DataEntry { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 MaxRecords { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.ADODBApi.Recordset DefaultRecordset { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.SchemaRowsources SchemaRowsources { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.SchemaRelationships SchemaRelationships { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api.PageRowsources PageRowsources { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.RecordsetDefs RecordsetDefs { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api.RecordsetDefs RootRecordsetDefs { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("OWC10", 1), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object PivotDefs { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string DefaultRecordsetName { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string XMLData { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.GroupLevels GroupLevels { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("OWC10", 1), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object Constants { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.ElementExtensions ElementExtensions { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool IsNew { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum RecordsetType { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.AllPageFields AllPageFields { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Section CurrentSection { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new NetOffice.OWC10Api.Enums.ProviderType ProviderType { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.AllGroupingDefs AllGroupingDefs { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool DisplayAlerts { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.DataPages DataPages { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 GridX { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 GridY { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 LoadError { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Enums.DefaultControlTypeEnum DefaultControlType { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool IsDirty { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool Busy { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string Version { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 MajorVersion { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string MinorVersion { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string BuildNumber { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string RevisionNumber { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool IsDataModelDirty { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Enums.DscOfflineTypeEnum OfflineType { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string OfflinePublication { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool Offline { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string OfflineSource { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Enums.DscXMLLocationEnum XMLLocation { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool UseXMLData { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string XMLDataTarget { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string ConnectionFile { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string DefaultRecordsetDefName { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string ConnectionStringFullPath { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api.SchemaDiagrams SchemaDiagrams { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string DBNSOwnerName { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        /// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
        /// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum FetchType = 2</param>
        [SupportByVersion("OWC10", 1)]
        NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption, object fetchType);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        NetOffice.ADODBApi.Recordset Execute(string recordsetName);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        /// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="dataAssistant">object dataAssistant</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SetDataAssistant(object dataAssistant);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="advise">object advise</param>
        /// <param name="sinkName">string sinkName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void DesignAdvise(object advise, string sinkName);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="sinkName">string sinkName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void DesignUnAdvise(string sinkName);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="pUnknownDropGoo">object pUnknownDropGoo</param>
        /// <param name="bstrRecordSetDefName">string bstrRecordSetDefName</param>
        /// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
        /// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
        /// <param name="pageRowsource">string pageRowsource</param>
        /// <param name="schemaRelationship">string schemaRelationship</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void ProcessDrop(object pUnknownDropGoo, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="rowsources">object rowsources</param>
        /// <param name="relationships">object relationships</param>
        /// <param name="fields">object fields</param>
        /// <param name="bstrRecordSetDefName">string bstrRecordSetDefName</param>
        /// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
        /// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
        /// <param name="pageRowsource">string pageRowsource</param>
        /// <param name="schemaRelationship">string schemaRelationship</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void ScriptDrop(object rowsources, object relationships, object fields, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="element">object element</param>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Section GetContainingSection(object element);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="rowsources">object rowsources</param>
        /// <param name="relationships">object relationships</param>
        /// <param name="fields">object fields</param>
        /// <param name="recordsetDef">string recordsetDef</param>
        /// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
        /// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
        /// <param name="dropRowsource">string dropRowsource</param>
        /// <param name="rowsourcesOut">object rowsourcesOut</param>
        /// <param name="relationshipsOut">object relationshipsOut</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void ScriptValidate(object rowsources, object relationships, object fields, string recordsetDef, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="unknownDropGoo">object unknownDropGoo</param>
        /// <param name="recordSetDefName">string recordSetDefName</param>
        /// <param name="location">NetOffice.OWC10Api.Enums.DscDropLocationEnum location</param>
        /// <param name="type">NetOffice.OWC10Api.Enums.DscDropTypeEnum type</param>
        /// <param name="dropRowsource">string dropRowsource</param>
        /// <param name="rowsourcesOut">object rowsourcesOut</param>
        /// <param name="relationshipsOut">object relationshipsOut</param>
        /// <param name="numberOfDrops">Int32 numberOfDrops</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void ValidateDrop(object unknownDropGoo, string recordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum location, NetOffice.OWC10Api.Enums.DscDropTypeEnum type, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut, out Int32 numberOfDrops);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="hyperlink">object hyperlink</param>
        /// <param name="part">NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum part</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        string HyperlinkPart(object hyperlink, NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum part);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SchemaRefresh();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="oldID">string oldID</param>
        /// <param name="newID">string newID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void UpdateElementID(string oldID, string newID);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void Reset();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lIndex">Int32 lIndex</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        string getDataMemberName(Int32 lIndex);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        Int32 getDataMemberCount();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="sectionElement">object sectionElement</param>
        /// <param name="recordSource">string recordSource</param>
        /// <param name="sectionType">NetOffice.OWC10Api.Enums.SectTypeEnum sectionType</param>
        /// <param name="groupLevel">NetOffice.OWC10Api.GroupLevel groupLevel</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void GetSectionInfo(object sectionElement, out string recordSource, out NetOffice.OWC10Api.Enums.SectTypeEnum sectionType, out NetOffice.OWC10Api.GroupLevel groupLevel);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordSource">string recordSource</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void DeleteRecordSourceIfUnused(string recordSource);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordSource">string recordSource</param>
        /// <param name="pageField">string pageField</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void DeletePageFieldIfUnused(string recordSource, string pageField);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="bstrRecordset">string bstrRecordset</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void ResetRecordset(string bstrRecordset);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="exportType">NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType</param>
        /// <param name="connectString">string connectString</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void GetExportableConnectString(NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType, out string connectString);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="eEncoding">optional NetOffice.OWC10Api.Enums.DscEncodingEnum eEncoding = 0</param>
        [SupportByVersion("OWC10", 1)]
        void ExportXML(object eEncoding);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void ExportXML();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        /// <param name="recordset">NetOffice.ADODBApi.Recordset recordset</param>
        [SupportByVersion("OWC10", 1)]
        void SetRootRecordset(string recordsetName, NetOffice.ADODBApi.Recordset recordset);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="onlineServer">string onlineServer</param>
        /// <param name="onlineDatabase">string onlineDatabase</param>
        /// <param name="offlineServer">string offlineServer</param>
        /// <param name="offlineDatabase">string offlineDatabase</param>
        [SupportByVersion("OWC10", 1)]
        void GetOfflineDisplayInfo(out string onlineServer, out string onlineDatabase, out string offlineServer, out string offlineDatabase);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="refreshType">optional NetOffice.OWC10Api.Enums.RefreshType RefreshType = 1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void Refresh(object refreshType);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Refresh();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
        /// <param name="fChild">Int32 fChild</param>
        /// <param name="ppGrouplevel">NetOffice.OWC10Api.GroupLevel ppGrouplevel</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void FindRelatedGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, Int32 fChild, out NetOffice.OWC10Api.GroupLevel ppGrouplevel);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="notification">NetOffice.OWC10Api.Enums.NotificationType notification</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void DllNotification(NetOffice.OWC10Api.Enums.NotificationType notification);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="suspend">bool suspend</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SuspendUndo(bool suspend);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void UpdateFocus();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="connectionString">string connectionString</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        bool IsValidDAPProvider(string connectionString);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        /// <param name="fullPrecision">optional object fullPrecision</param>
        /// <param name="triangulationPrecision">optional object triangulationPrecision</param>
        [SupportByVersion("OWC10", 1)]
        Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        Double EuroConvert(Double number, string sourceCurrency, string targetCurrency);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        /// <param name="fullPrecision">optional object fullPrecision</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        String[] GetDAPProviders();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="synchronizing">bool synchronizing</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SetSynchronizing(bool synchronizing);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="displayError">bool displayError</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SetDisplayError(bool displayError);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="suspend">bool suspend</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SuspendXMLReExecute(bool suspend);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="firePropChange">bool firePropChange</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void SetFirePropChange(bool firePropChange);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="value">object value</param>
        /// <param name="valueIfNull">optional object valueIfNull</param>
        [SupportByVersion("OWC10", 1)]
        object Nz(object value, object valueIfNull);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="value">object value</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        object Nz(object value);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void RefreshJetCache();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void AutoRefreshOfflineSource();

        #endregion
    }
}
