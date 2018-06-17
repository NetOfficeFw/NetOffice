using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _WizHook 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("CB9D3171-4728-11D1-8334-006008197CC8")]
    [CoClassSource(typeof(NetOffice.AccessApi.WizHook))]
    public interface _WizHook : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 Key { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VBIDEApi._VBProject DbcVbProject { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="bstrConnectionString">string bstrConnectionString</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool get_IsMatchToDbcConnectString(string bstrConnectionString);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_IsMatchToDbcConnectString
		/// </summary>
		/// <param name="bstrConnectionString">string bstrConnectionString</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_IsMatchToDbcConnectString")]
		bool IsMatchToDbcConnectString(string bstrConnectionString);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="actid">Int32 actid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string NameFromActid(Int32 actid);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="actid">Int32 actid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 ArgsOfActid(Int32 actid);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="script">string script</param>
		/// <param name="label">string label</param>
		/// <param name="openMode">Int32 openMode</param>
		/// <param name="extra">Int32 extra</param>
		/// <param name="version">Int32 version</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 OpenScript(string script, string label, Int32 openMode, Int32 extra, Int32 version);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hScr">Int32 hScr</param>
		/// <param name="scriptColumn">Int32 scriptColumn</param>
		/// <param name="value">string value</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool GetScriptString(Int32 hScr, Int32 scriptColumn, string value);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hScr">Int32 hScr</param>
		/// <param name="scriptColumn">Int32 scriptColumn</param>
		/// <param name="value">string value</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool SaveScriptString(Int32 hScr, Int32 scriptColumn, string value);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool GlobalProcExists(string name);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="table">string table</param>
		/// <param name="columns">string columns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool TableFieldHasUniqueIndex(string table, string columns);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="flags">Int32 flags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool BracketString(string _string, Int32 flags);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="helpFile">string helpFile</param>
		/// <param name="wCmd">Int32 wCmd</param>
		/// <param name="contextID">Int32 contextID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool WizHelp(string helpFile, Int32 wCmd, Int32 contextID);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="cancelled">bool cancelled</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool OpenPictureFile(string file, bool cancelled);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_in">string in</param>
		/// <param name="_out">string out</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool EnglishPictToLocal(string _in, string _out);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_in">string in</param>
		/// <param name="_out">string out</param>
		/// <param name="parseFlags">Int32 parseFlags</param>
		/// <param name="translateFlags">Int32 translateFlags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool TranslateExpression(string _in, string _out, Int32 parseFlags, Int32 translateFlags);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool FileExists(string file);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="relativePath">string relativePath</param>
		/// <param name="fullPath">string fullPath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FullPath(string relativePath, string fullPath);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="drive">string drive</param>
		/// <param name="dir">string dir</param>
		/// <param name="file">string file</param>
		/// <param name="ext">string ext</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SplitPath(string path, string drive, string dir, string file, string ext);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fontName">string fontName</param>
		/// <param name="size">Int32 size</param>
		/// <param name="weight">Int32 weight</param>
		/// <param name="italic">bool italic</param>
		/// <param name="underline">bool underline</param>
		/// <param name="cch">Int32 cch</param>
		/// <param name="caption">string caption</param>
		/// <param name="maxWidthCch">Int32 maxWidthCch</param>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool TwipsFromFont(string fontName, Int32 size, Int32 weight, bool italic, bool underline, Int32 cch, string caption, Int32 maxWidthCch, Int32 dx, Int32 dy);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recordSource">string recordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 ObjTypOfRecordSource(string recordSource);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="identifier">string identifier</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool IsValidIdent(string identifier);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="array">String[] array</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SortStringArray(String[] array);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="workspace">NetOffice.DAOApi.Workspace workspace</param>
		/// <param name="database">NetOffice.DAOApi.Database database</param>
		/// <param name="table">string table</param>
		/// <param name="returnDebugInfo">bool returnDebugInfo</param>
		/// <param name="results">string results</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 AnalyzeTable(NetOffice.DAOApi.Workspace workspace, NetOffice.DAOApi.Database database, string table, bool returnDebugInfo, string results);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="workspace">NetOffice.DAOApi.Workspace workspace</param>
		/// <param name="database">NetOffice.DAOApi.Database database</param>
		/// <param name="query">string query</param>
		/// <param name="results">string results</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 AnalyzeQuery(NetOffice.DAOApi.Workspace workspace, NetOffice.DAOApi.Database database, string query, string results);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwndOwner">Int32 hwndOwner</param>
		/// <param name="appName">string appName</param>
		/// <param name="dlgTitle">string dlgTitle</param>
		/// <param name="openTitle">string openTitle</param>
		/// <param name="file">string file</param>
		/// <param name="initialDir">string initialDir</param>
		/// <param name="filter">string filter</param>
		/// <param name="filterIndex">Int32 filterIndex</param>
		/// <param name="view">Int32 view</param>
		/// <param name="flags">Int32 flags</param>
		/// <param name="fOpen">bool fOpen</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 GetFileName(Int32 hwndOwner, string appName, string dlgTitle, string openTitle, string file, string initialDir, string filter, Int32 filterIndex, Int32 view, Int32 flags, bool fOpen);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dpName">string dpName</param>
		/// <param name="ctlName">string ctlName</param>
		/// <param name="typ">Int32 typ</param>
		/// <param name="section">string section</param>
		/// <param name="sectionType">Int32 sectionType</param>
		/// <param name="appletCode">string appletCode</param>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CreateDataPageControl(string dpName, string ctlName, Int32 typ, string section, Int32 sectionType, string appletCode, Int32 x, Int32 y, Int32 dx, Int32 dy);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fStart">bool fStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void KnownWizLeaks(bool fStart);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrDbName">string bstrDbName</param>
		/// <param name="bstrConnect">string bstrConnect</param>
		/// <param name="bstrPasswd">string bstrPasswd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool SetVbaPassword(string bstrDbName, string bstrConnect, string bstrPasswd);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string LocalFont();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="objtyp">Int16 objtyp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SaveObject(string bstrName, Int16 objtyp);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 CurrentLangID();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 KeyboardLangID();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string AccessUserDataDir();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OfficeAddInDir();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dpName">string dpName</param>
		/// <param name="fileToInsert">string fileToInsert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string EmbedFileOnDataPage(string dpName, string fileToInsert);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fRptToFile">bool fRptToFile</param>
		/// <param name="bstrFileOut">string bstrFileOut</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ReportLeaksToFile(bool fRptToFile, string bstrFileOut);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrFilename">string bstrFilename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void LoadImexSpecSolution(string bstrFilename);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fBlockKeys">bool fBlockKeys</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetDpBlockKeyInput(bool fBlockKeys);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="objType">NetOffice.AccessApi.Enums.AcObjectType objType</param>
		/// <param name="attribs">Int32 attribs</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool FirstDbcDataObject(string name, NetOffice.AccessApi.Enums.AcObjectType objType, Int32 attribs);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool CloseCurrentDatabase();

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrWhich">string bstrWhich</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string AccessWizFilePath(string bstrWhich);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool HideDates();

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrBase">string bstrBase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string GetColumns(string bstrBase);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExt">string bstrExt</param>
		/// <param name="bstrFilename">string bstrFilename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 GetFileOdso(string bstrExt, string bstrFilename);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrBase">string bstrBase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string GetInfoForColumns(string bstrBase);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwndOwner">Int32 hwndOwner</param>
		/// <param name="appName">string appName</param>
		/// <param name="dlgTitle">string dlgTitle</param>
		/// <param name="openTitle">string openTitle</param>
		/// <param name="file">string file</param>
		/// <param name="initialDir">string initialDir</param>
		/// <param name="filter">string filter</param>
		/// <param name="filterIndex">Int32 filterIndex</param>
		/// <param name="view">Int32 view</param>
		/// <param name="flags">Int32 flags</param>
		/// <param name="fOpen">bool fOpen</param>
		/// <param name="fFileSystem">object fFileSystem</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 GetFileName2(Int32 hwndOwner, string appName, string dlgTitle, string openTitle, string file, string initialDir, string filter, Int32 filterIndex, Int32 view, Int32 flags, bool fOpen, object fFileSystem);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fBlockKeys">bool fBlockKeys</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool FGetMSDE(bool fBlockKeys);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="wStyle">Int32 wStyle</param>
		/// <param name="idHelpID">Int32 idHelpID</param>
		/// <param name="bstrHelpFileName">string bstrHelpFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 WizMsgBox(string bstrText, string bstrCaption, Int32 wStyle, Int32 idHelpID, string bstrHelpFileName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pbstrUID">string pbstrUID</param>
		/// <param name="pbstrPwd">string pbstrPwd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool AdpUIDPwd(string pbstrUID, string pbstrPwd);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="lWhich">Int32 lWhich</param>
		/// <param name="vValue">object vValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void SetWizGlob(Int32 lWhich, object vValue);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="lWhich">Int32 lWhich</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		object GetWizGlob(Int32 lWhich);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrADPName">string bstrADPName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void WizCopyCmdbars(string bstrADPName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 GetCurrentView(string bstrTableName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wch">Int32 wch</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool FIsFEWch(Int32 wch);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		string GetAccWizRCPath();

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objtyp">Int16 objtyp</param>
		/// <param name="bstrObjName">string bstrObjName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool FCreateNameMap(Int16 objtyp, string bstrObjName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		string GetAdeRegistryPath();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrSpecXML">string bstrSpecXML</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void ExecuteTempImexSpec(string bstrSpecXML);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		bool FCacheStatus();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrStatus">string bstrStatus</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void CacheStatus(string bstrStatus);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrSpecName">string bstrSpecName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void SetDefaultSpecName(string bstrSpecName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		string GetImexTblName();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrPropertyName">string bstrPropertyName</param>
		/// <param name="fServer">bool fServer</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		string GetLinkedListProperty(string bstrTableName, string bstrPropertyName, bool fServer);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="pProperty">NetOffice.AccessApi._AccessProperty pProperty</param>
		/// <param name="openMode">Int32 openMode</param>
		/// <param name="extra">Int32 extra</param>
		/// <param name="version">Int32 version</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 OpenEmScript(NetOffice.AccessApi._AccessProperty pProperty, Int32 openMode, Int32 extra, Int32 version);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		string GetDisabledExtensions();

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		/// <param name="iobjtyp">NetOffice.AccessApi.Enums.AcObjectType iobjtyp</param>
		/// <param name="fTablesAsClient">bool fTablesAsClient</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		Int32 GetObjPubOption(string bstrObjectName, NetOffice.AccessApi.Enums.AcObjectType iobjtyp, bool fTablesAsClient);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		bool FIsPublishedXasTable(string bstrObjectName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		bool FIsXasDb();

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		/// <param name="iobjtyp">NetOffice.AccessApi.Enums.AcObjectType iobjtyp</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		bool FIsValidXasObjectName(string bstrObjectName, NetOffice.AccessApi.Enums.AcObjectType iobjtyp);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrObjectName">string bstrObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object LoadResourceLibrary(string bstrObjectName);

		#endregion
	}
}
