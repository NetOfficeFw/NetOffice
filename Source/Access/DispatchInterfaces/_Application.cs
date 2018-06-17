using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Application
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("68CCE6C0-6129-101B-AF4E-00AA003F0F07")]
    [CoClassSource(typeof(NetOffice.AccessApi.Application))]
	public interface _Application : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192087.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836400.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822407.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object CodeContextObject { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835352.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string MenuBar { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845319.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 CurrentObjectType { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196795.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string CurrentObjectName { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837183.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Forms Forms { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834339.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Reports Reports { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835056.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Screen Screen { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845564.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.DoCmd DoCmd { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195236.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ShortcutMenuBar { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821493.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836033.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool UserControl { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821724.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.DAOApi.DBEngine DBEngine { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821379.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.CommandBars CommandBars { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Assistant Assistant { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835326.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.References References { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836265.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Modules Modules { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.FileSearch FileSearch { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823044.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool IsCompiled { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822476.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.VBIDEApi.VBE VBE { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.DataAccessPages DataAccessPages { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string ADOConnectString { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193770.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.CurrentProject CurrentProject { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193230.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.CurrentData CurrentData { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197047.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.CodeProject CodeProject { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836912.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.CodeData CodeData { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.AccessApi.WizHook WizHook { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822077.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ProductCode { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822463.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.COMAddIns COMAddIns { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194961.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.DefaultWebOptions DefaultWebOptions { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836634.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.LanguageSettings LanguageSettings { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.AnswerWizard AnswerWizard { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822721.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object VGXFrameInterval { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx </remarks>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx </remarks>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersion("Access", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845884.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool BrokenReference { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195779.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Printers Printers { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821394.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._Printer Printer { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192859.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835096.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 Build { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191715.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.NewFile NewFileTaskPane { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845345.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._AutoCorrect AutoCorrect { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193178.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845034.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.MacroError MacroError { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192459.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.TempVars TempVars { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192450.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.OfficeApi.IAssistance Assistance { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837286.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		NetOffice.AccessApi.WebServices WebServices { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.AccessApi.LocalVars LocalVars { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj249062.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		NetOffice.AccessApi.ReturnVars ReturnVars { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void NewCurrentDatabase(string filepath);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object template</param>
		/// <param name="siteAddress">optional string SiteAddress = </param>
		/// <param name="listID">optional string ListID = </param>
		[SupportByVersion("Access", 12,14,15,16)]
		void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress, object listID);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void NewCurrentDatabase(string filepath, object fileFormat);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object template</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void NewCurrentDatabase(string filepath, object fileFormat, object template);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object template</param>
		/// <param name="siteAddress">optional string SiteAddress = </param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenCurrentDatabase(string filepath, object exclusive);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		/// <param name="bstrPassword">optional string bstrPassword = </param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenCurrentDatabase(string filepath, object exclusive, object bstrPassword);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenCurrentDatabase(string filepath);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192308.aspx </remarks>
		/// <param name="optionName">string optionName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object GetOption(string optionName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195513.aspx </remarks>
		/// <param name="optionName">string optionName</param>
		/// <param name="setting">object setting</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetOption(string optionName, object setting);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx </remarks>
		/// <param name="echoOn">Int16 echoOn</param>
		/// <param name="bstrStatusBarText">optional string bstrStatusBarText = </param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Echo(Int16 echoOn, object bstrStatusBarText);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx </remarks>
		/// <param name="echoOn">Int16 echoOn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Echo(Int16 echoOn);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836850.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CloseCurrentDatabase();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx </remarks>
		/// <param name="option">optional NetOffice.AccessApi.Enums.AcQuitOption Option = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Quit(object option);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		/// <param name="argument2">optional object argument2</param>
		/// <param name="argument3">optional object argument3</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2, object argument3);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		/// <param name="argument2">optional object argument2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
		/// <param name="database">optional object database</param>
		/// <param name="formTemplate">optional object formTemplate</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Form CreateForm(object database, object formTemplate);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Form CreateForm();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
		/// <param name="database">optional object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Form CreateForm(object database);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
		/// <param name="database">optional object database</param>
		/// <param name="reportTemplate">optional object reportTemplate</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Report CreateReport(object database, object reportTemplate);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Report CreateReport();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
		/// <param name="database">optional object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Report CreateReport(object database);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlSource">string controlSource</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlEx(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlName">string controlName</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlEx(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836733.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlName">string controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DeleteControl(string formName, string controlName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191904.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlName">string controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DeleteReportControl(string reportName, string controlName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197044.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="expression">string expression</param>
		/// <param name="header">Int16 header</param>
		/// <param name="footer">Int16 footer</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 CreateGroupLevel(string reportName, string expression, Int16 header, Int16 footer);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DMin(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DMin(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DMax(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DMax(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DSum(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DSum(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DAvg(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DAvg(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DLookup(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DLookup(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DLast(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DLast(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DVar(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DVar(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DVarP(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DVarP(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DStDev(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DStDev(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DStDevP(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DStDevP(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DFirst(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DFirst(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DCount(string expr, string domain, object criteria);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DCount(string expr, string domain);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834705.aspx </remarks>
		/// <param name="stringExpr">string stringExpr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Eval(string stringExpr);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845778.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string CurrentUser();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196189.aspx </remarks>
		/// <param name="application">string application</param>
		/// <param name="topic">string topic</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object DDEInitiate(string application, string topic);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197936.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="command">string command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DDEExecute(object chanNum, string command);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194752.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="item">string item</param>
		/// <param name="data">string data</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DDEPoke(object chanNum, string item, string data);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823145.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="item">string item</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string DDERequest(object chanNum, string item);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197795.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DDETerminate(object chanNum);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845193.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DDETerminateAll();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835631.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.DAOApi.Database CurrentDb();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196457.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.DAOApi.Database CodeDb();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwnd">Int32 hwnd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void BeginUndoable(Int32 hwnd);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="yesno">Int16 yesno</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetUndoRecording(Int16 yesno);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845070.aspx </remarks>
		/// <param name="field">string field</param>
		/// <param name="fieldType">Int16 fieldType</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string BuildCriteria(string field, Int16 fieldType, string expression);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="moduleName">string moduleName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void InsertText(string text, string moduleName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ReloadAddIns();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836901.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.DAOApi.Workspace DefaultWorkspaceClone();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197957.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RefreshTitleBar();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="changeFrom">string changeFrom</param>
		/// <param name="changeTo">string changeTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void AddAutoCorrect(string changeFrom, string changeTo);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="changeFrom">string changeFrom</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DelAutoCorrect(string changeFrom);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196179.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 hWndAccessApp();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		/// <param name="arg30">optional object arg30</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx </remarks>
		/// <param name="value">object value</param>
		/// <param name="valueIfNull">optional object valueIfNull</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Nz(object value, object valueIfNull);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx </remarks>
		/// <param name="value">object value</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Nz(object value);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835072.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object LoadPicture(string fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objtyp">Int32 objtyp</param>
		/// <param name="moduleName">string moduleName</param>
		/// <param name="fileName">string fileName</param>
		/// <param name="token">Int32 token</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ReplaceModule(Int32 objtyp, string moduleName, string fileName, Int32 token);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196488.aspx </remarks>
		/// <param name="errorNumber">object errorNumber</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object AccessError(object errorNumber);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object BuilderString();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193935.aspx </remarks>
		/// <param name="guid">object guid</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object StringFromGUID(object guid);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197675.aspx </remarks>
		/// <param name="_string">object string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object GUIDFromString(object _string);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="id">Int32 id</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object AppLoadString(Int32 id);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SaveAsText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void LoadFromText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823011.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void AddToFavorites();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194960.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RefreshDatabaseWindow();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191909.aspx </remarks>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunCommand(NetOffice.AccessApi.Enums.AcCommand command);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx </remarks>
		/// <param name="hyperlink">object hyperlink</param>
		/// <param name="part">optional NetOffice.AccessApi.Enums.AcHyperlinkPart Part = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string HyperlinkPart(object hyperlink, object part);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx </remarks>
		/// <param name="hyperlink">object hyperlink</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string HyperlinkPart(object hyperlink);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821756.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool GetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822459.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fHidden">bool fHidden</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, bool fHidden);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="createNewFile">optional bool CreateNewFile = true</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName, object createNewFile);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.DataAccessPage CreateDataAccessPage();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void NewAccessProject(string filepath, object connect);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void NewAccessProject(string filepath);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenAccessProject(string filepath, object exclusive);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenAccessProject(string filepath);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CreateAccessProject(string filepath, object connect);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CreateAccessProject(string filepath);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		/// <param name="fullPrecision">optional object fullPrecision</param>
		/// <param name="triangulationPrecision">optional object triangulationPrecision</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Double EuroConvert(Double number, string sourceCurrency, string targetCurrency);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		/// <param name="fullPrecision">optional object fullPrecision</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenCurrentDatabaseOld(string filepath, object exclusive);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenCurrentDatabaseOld(string filepath);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		/// <param name="replace">optional bool Replace = false</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID, object replace);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CreateNewWorkgroupFile();

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CreateNewWorkgroupFile(object path);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CreateNewWorkgroupFile(object path, object name);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CreateNewWorkgroupFile(object path, object name, object company);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195103.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void SetDefaultWorkgroupFile(string path);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193465.aspx </remarks>
		/// <param name="sourceFilename">string sourceFilename</param>
		/// <param name="destinationFilename">string destinationFilename</param>
		/// <param name="destinationFileFormat">NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ConvertAccessProject(string sourceFilename, string destinationFilename, NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx </remarks>
		/// <param name="sourceFile">string sourceFile</param>
		/// <param name="destinationFile">string destinationFile</param>
		/// <param name="logFile">optional bool LogFile = false</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool CompactRepair(string sourceFile, string destinationFile, object logFile);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx </remarks>
		/// <param name="sourceFile">string sourceFile</param>
		/// <param name="destinationFile">string destinationFile</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool CompactRepair(string sourceFile, string destinationFile);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional Int32 OtherFlags = 0</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional NetOffice.AccessApi.Enums.AcExportXMLOtherFlags OtherFlags = 0</param>
		/// <param name="whereCondition">optional string WhereCondition = </param>
		/// <param name="additionalData">optional object additionalData</param>
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition, object additionalData);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional NetOffice.AccessApi.Enums.AcExportXMLOtherFlags OtherFlags = 0</param>
		/// <param name="whereCondition">optional string WhereCondition = </param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="importOptions">optional NetOffice.AccessApi.Enums.AcImportXMLOption ImportOptions = 1</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ImportXML(string dataSource, object importOptions);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void ImportXML(string dataSource);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional Int32 OtherFlags = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		/// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
		/// <param name="scriptOption">optional NetOffice.AccessApi.Enums.AcTransformXMLScriptOption ScriptOption = 1</param>
		[SupportByVersion("Access", 11,12,14,15,16)]
		void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput, object scriptOption);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void TransformXML(string dataSource, string transformSource, string outputTarget);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		/// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834773.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._AdditionalData CreateAdditionalData();

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void NewCurrentDatabaseOld(string filepath);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlSource">string controlSource</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateControlExOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlName">string controlName</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Control CreateReportControlExOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx </remarks>
		/// <param name="richText">object richText</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Access", 12,14,15,16)]
		string PlainText(object richText, object length);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx </remarks>
		/// <param name="richText">object richText</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		string PlainText(object richText);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx </remarks>
		/// <param name="plainText">object plainText</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Access", 12,14,15,16)]
		string HtmlEncode(object plainText, object length);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx </remarks>
		/// <param name="plainText">object plainText</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		string HtmlEncode(object plainText);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194416.aspx </remarks>
		/// <param name="customUIName">string customUIName</param>
		/// <param name="customUIXML">string customUIXML</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void LoadCustomUI(string customUIName, string customUIXML);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193467.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void ExportNavigationPane(string path);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fAppendOnly">optional bool fAppendOnly = false</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void ImportNavigationPane(string path, object fAppendOnly);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void ImportNavigationPane(string path);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835727.aspx </remarks>
		/// <param name="tableName">string tableName</param>
		/// <param name="columnName">string columnName</param>
		/// <param name="queryString">string queryString</param>
		[SupportByVersion("Access", 12,14,15,16)]
		string ColumnHistory(string tableName, string columnName, string queryString);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage, object toPage);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821429.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845765.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		void LoadFromAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		/// <param name="applicationPart">optional object applicationPart</param>
		/// <param name="includeData">optional object includeData</param>
		/// <param name="variation">optional object variation</param>
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData, object variation);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		/// <param name="applicationPart">optional object applicationPart</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		/// <param name="applicationPart">optional object applicationPart</param>
		/// <param name="includeData">optional object includeData</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835421.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 14,15,16)]
		void InstantiateTemplate(string path);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834388.aspx </remarks>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption</param>
		[SupportByVersion("Access", 14,15,16)]
		object CurrentWebUser(NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836539.aspx </remarks>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption</param>
		[SupportByVersion("Access", 14,15,16)]
		object CurrentWebUserGroups(NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193453.aspx </remarks>
		/// <param name="groupNameOrID">object groupNameOrID</param>
		[SupportByVersion("Access", 14,15,16)]
		bool IsCurrentWebUserInGroup(object groupNameOrID);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834368.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 14,15,16)]
		void DirtyObject(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		bool IsClient();

		#endregion
	}
}
