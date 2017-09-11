using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Access, 9,10,11,12,14,15,16
    ///</summary>
    [SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(AccessApi.Application))]
	public static class GlobalModule
	{
		#region Fields

		private static ICOMObject _instance;
        
        #endregion

        #region Internal Properties

        internal static ICOMObject Instance
        {
            get
            {
                return _instance;
            }
            set
            {
                if ((null == value) || (null == _instance))
                    _instance = value;
            }
        }

        internal static Core Factory
		{
			get
			{
				if(null != _instance)
					 return _instance.Factory;
			else
				return Core.Default;
			}
		}

		internal static Invoker Invoker
		{
			get
			{
				if(null != _instance)
					 return _instance.Invoker;
			else
				return Invoker.Default;
			}
		}

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192087.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Application Application
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(_instance, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836400.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public static object Parent
		{
			get
			{
                return Factory.ExecuteReferencePropertyGet(_instance, "Parent");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822407.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public static object CodeContextObject
		{
			get
			{
                return Factory.ExecuteReferencePropertyGet(_instance, "CodeContextObject");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835352.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string MenuBar
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "MenuBar");
			}
			set
			{
                Factory.ExecuteValuePropertySet(_instance, "MenuBar", value);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845319.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Int32 CurrentObjectType
		{
			get
			{
                return Factory.ExecuteInt32PropertyGet(_instance, "CurrentObjectType");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196795.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string CurrentObjectName
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "CurrentObjectName");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837183.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Forms Forms
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Forms>(_instance, "Forms", NetOffice.AccessApi.Forms.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834339.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Reports Reports
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Reports>(_instance, "Reports", NetOffice.AccessApi.Reports.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835056.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Screen Screen
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Screen>(_instance, "Screen", NetOffice.AccessApi.Screen.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845564.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DoCmd DoCmd
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.DoCmd>(_instance, "DoCmd", NetOffice.AccessApi.DoCmd.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195236.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string ShortcutMenuBar
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "ShortcutMenuBar");
			}
			set
			{
                Factory.ExecuteValuePropertySet(_instance, "ShortcutMenuBar", value);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821493.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool Visible
		{
			get
			{
                return Factory.ExecuteBoolPropertyGet(_instance, "Visible");
			}
			set
			{
                Factory.ExecuteValuePropertySet(_instance, "Visible", value);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836033.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool UserControl
		{
			get
			{
                return Factory.ExecuteBoolPropertyGet(_instance, "UserControl");
			}
			set
			{
                Factory.ExecuteValuePropertySet(_instance, "UserControl", value);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821724.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.DAOApi.DBEngine DBEngine
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.DBEngine>(_instance, "DBEngine", NetOffice.DAOApi.DBEngine.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821379.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(_instance, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(_instance, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835326.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.References References
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.References>(_instance, "References", NetOffice.AccessApi.References.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836265.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Modules Modules
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Modules>(_instance, "Modules", NetOffice.AccessApi.Modules.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(_instance, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823044.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool IsCompiled
		{
			get
			{
                return Factory.ExecuteBoolPropertyGet(_instance, "IsCompiled");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get 
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822476.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(_instance, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DataAccessPages DataAccessPages
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.DataAccessPages>(_instance, "DataAccessPages", NetOffice.AccessApi.DataAccessPages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static string ADOConnectString
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "ADOConnectString");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193770.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CurrentProject CurrentProject
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CurrentProject>(_instance, "CurrentProject", NetOffice.AccessApi.CurrentProject.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193230.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CurrentData CurrentData
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CurrentData>(_instance, "CurrentData", NetOffice.AccessApi.CurrentData.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197047.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CodeProject CodeProject
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CodeProject>(_instance, "CodeProject", NetOffice.AccessApi.CodeProject.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836912.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CodeData CodeData
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CodeData>(_instance, "CodeData", NetOffice.AccessApi.CodeData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.AccessApi.WizHook WizHook
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.WizHook>(_instance, "WizHook", NetOffice.AccessApi.WizHook.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822077.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string ProductCode
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "ProductCode");
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822463.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(_instance, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194961.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string Name
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DefaultWebOptions DefaultWebOptions
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.DefaultWebOptions>(_instance, "DefaultWebOptions", NetOffice.AccessApi.DefaultWebOptions.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836634.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(_instance, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(_instance, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822721.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFeatureInstall>(_instance, "FeatureInstall");
			}
			set
			{
                Factory.ExecuteEnumPropertySet(_instance, "FeatureInstall", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static object VGXFrameInterval
		{
			get
			{
                return Factory.ExecuteVariantPropertyGet(_instance, "VGXFrameInterval");
			}
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx </remarks>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(_instance, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, dialogType);
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Alias for get_FileDialog
        /// </summary>
        /// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx </remarks>
        [SupportByVersion("Access", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		public static NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return get_FileDialog(dialogType);
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845884.aspx </remarks>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		public static bool BrokenReference
		{
			get
			{
                return Factory.ExecuteBoolPropertyGet(_instance, "BrokenReference");
			}
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195779.aspx </remarks>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Printers Printers
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Printers>(_instance, "Printers", NetOffice.AccessApi.Printers.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821394.aspx </remarks>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.AccessApi._Printer Printer
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Printer>(_instance, "Printer");
            }
            set
            {
                Factory.ExecuteReferencePropertySet(_instance, "Printer", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(_instance, "MsoDebugOptions", NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192859.aspx </remarks>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		public static string Version
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "Version");
			}
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835096.aspx </remarks>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		public static Int32 Build
		{
			get
			{
                return Factory.ExecuteInt32PropertyGet(_instance, "Build");
			}
		}

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191715.aspx </remarks>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.NewFile NewFileTaskPane
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(_instance, "NewFileTaskPane", NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845345.aspx </remarks>
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.AccessApi._AutoCorrect AutoCorrect
		{
			get
			{
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._AutoCorrect>(_instance, "AutoCorrect");
            }
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193178.aspx </remarks>
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutomationSecurity>(_instance, "AutomationSecurity");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "AutomationSecurity", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845034.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.MacroError MacroError
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.MacroError>(_instance, "MacroError", NetOffice.AccessApi.MacroError.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192459.aspx </remarks>
        [SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.TempVars TempVars
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.TempVars>(_instance, "TempVars", NetOffice.AccessApi.TempVars.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192450.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(_instance, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837286.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public static NetOffice.AccessApi.WebServices WebServices
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.WebServices>(_instance, "WebServices", NetOffice.AccessApi.WebServices.LateBindingApiWrapperType);
            }
        }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.AccessApi.LocalVars LocalVars
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.LocalVars>(_instance, "LocalVars", NetOffice.AccessApi.LocalVars.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj249062.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
		public static NetOffice.AccessApi.ReturnVars ReturnVars
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.ReturnVars>(_instance, "ReturnVars", NetOffice.AccessApi.ReturnVars.LateBindingApiWrapperType);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void NewCurrentDatabase(string filepath)
		{
            Factory.ExecuteMethod(_instance, "NewCurrentDatabase", filepath);
        }

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
		public static void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress, object listID)
		{
            Factory.ExecuteMethod(_instance, "NewCurrentDatabase", new object[] { filepath, fileFormat, template, siteAddress, listID });
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
        [CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabase(string filepath, object fileFormat)
		{
            Factory.ExecuteMethod(_instance, "NewCurrentDatabase", filepath, fileFormat);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
        /// <param name="template">optional object template</param>
        [CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabase(string filepath, object fileFormat, object template)
		{
            Factory.ExecuteMethod(_instance, "NewCurrentDatabase", filepath, fileFormat, template);
        }

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
		public static void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress)
		{
            Factory.ExecuteMethod(_instance, "NewCurrentDatabase", filepath, fileFormat, template, siteAddress);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="exclusive">optional bool Exclusive = false</param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
        public static void OpenCurrentDatabase(string filepath, object exclusive)
		{
            Factory.ExecuteMethod(_instance, "OpenCurrentDatabase", filepath, exclusive);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="exclusive">optional bool Exclusive = false</param>
        /// <param name="bstrPassword">optional string bstrPassword = </param>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void OpenCurrentDatabase(string filepath, object exclusive, object bstrPassword)
		{
            Factory.ExecuteMethod(_instance, "OpenCurrentDatabase", filepath, exclusive, bstrPassword);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        [CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void OpenCurrentDatabase(string filepath)
		{
            Factory.ExecuteMethod(_instance, "OpenCurrentDatabase", filepath);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192308.aspx </remarks>
        /// <param name="optionName">string optionName</param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object GetOption(string optionName)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "GetOption", optionName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195513.aspx </remarks>
        /// <param name="optionName">string optionName</param>
        /// <param name="setting">object setting</param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void SetOption(string optionName, object setting)
		{
            Factory.ExecuteMethod(_instance, "SetOption", optionName, setting);
        }

        /// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx </remarks>
		/// <param name="echoOn">Int16 echoOn</param>
		/// <param name="bstrStatusBarText">optional string bstrStatusBarText = </param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Echo(Int16 echoOn, object bstrStatusBarText)
		{
            Factory.ExecuteMethod(_instance, "Echo", echoOn, bstrStatusBarText);
        }
        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx </remarks>
        /// <param name="echoOn">Int16 echoOn</param>
        [CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Echo(Int16 echoOn)
		{
            Factory.ExecuteMethod(_instance, "Echo", echoOn);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836850.aspx </remarks>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void CloseCurrentDatabase()
		{
            Factory.ExecuteMethod(_instance, "CloseCurrentDatabase");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx </remarks>
        /// <param name="option">optional NetOffice.AccessApi.Enums.AcQuitOption Option = 1</param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Quit(object option)
		{
            Factory.ExecuteMethod(_instance, "Quit", option);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx </remarks>
        [CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Quit()
		{
            Factory.ExecuteMethod(_instance, "Quit");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
        /// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
        /// <param name="argument2">optional object argument2</param>
        /// <param name="argument3">optional object argument3</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2, object argument3)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "SysCmd", action, argument2, argument3);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
        /// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "SysCmd", action);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
        /// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
        /// <param name="argument2">optional object argument2</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "SysCmd", action, argument2);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
        /// <param name="database">optional object database</param>
        /// <param name="formTemplate">optional object formTemplate</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Form CreateForm(object database, object formTemplate)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Form>(_instance, "CreateForm", NetOffice.AccessApi.Form.LateBindingApiWrapperType, database, formTemplate);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Form CreateForm()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Form>(_instance, "CreateForm", NetOffice.AccessApi.Form.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
        /// <param name="database">optional object database</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Form CreateForm(object database)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Form>(_instance, "CreateForm", NetOffice.AccessApi.Form.LateBindingApiWrapperType, database);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
        /// <param name="database">optional object database</param>
        /// <param name="reportTemplate">optional object reportTemplate</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Report CreateReport(object database, object reportTemplate)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Report>(_instance, "CreateReport", NetOffice.AccessApi.Report.LateBindingApiWrapperType, database, reportTemplate);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Report CreateReport()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Report>(_instance, "CreateReport", NetOffice.AccessApi.Report.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
        /// <param name="database">optional object database</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Report CreateReport(object database)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Report>(_instance, "CreateReport", NetOffice.AccessApi.Report.LateBindingApiWrapperType, database);
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left, top, width, height });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
        /// <param name="formName">string formName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, formName, controlType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
        /// <param name="formName">string formName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, formName, controlType, section);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
        /// <param name="formName">string formName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        /// <param name="parent">optional object parent</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, formName, controlType, section, parent);
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left, top });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left, top, width });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left, top, width, height });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, reportName, controlType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, reportName, controlType, section);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        /// <param name="parent">optional object parent</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, reportName, controlType, section, parent);
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left, top });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left, top, width });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlEx(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlEx", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, controlSource, left, top, width, height });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlEx(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlEx", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, controlName, left, top, width, height });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836733.aspx </remarks>
        /// <param name="formName">string formName</param>
        /// <param name="controlName">string controlName</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DeleteControl(string formName, string controlName)
        {
            Factory.ExecuteMethod(_instance, "DeleteControl", formName, controlName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191904.aspx </remarks>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlName">string controlName</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DeleteReportControl(string reportName, string controlName)
        {
            Factory.ExecuteMethod(_instance, "DeleteReportControl", reportName, controlName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197044.aspx </remarks>
        /// <param name="reportName">string reportName</param>
        /// <param name="expression">string expression</param>
        /// <param name="header">Int16 header</param>
        /// <param name="footer">Int16 footer</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 CreateGroupLevel(string reportName, string expression, Int16 header, Int16 footer)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "CreateGroupLevel", reportName, expression, header, footer);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DMin(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DMin", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DMin(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DMin", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DMax(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DMax", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DMax(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DMax", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DSum(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DSum", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DSum(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DSum", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DAvg(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DAvg", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DAvg(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DAvg", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DLookup(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DLookup", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DLookup(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DLookup", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DLast(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DLast", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DLast(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DLast", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DVar(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DVar", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DVar(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DVar", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DVarP(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DVarP", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DVarP(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DVarP", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DStDev(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DStDev", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DStDev(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DStDev", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DStDevP(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DStDevP", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DStDevP(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DStDevP", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DFirst(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DFirst", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DFirst(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DFirst", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        /// <param name="criteria">optional object criteria</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DCount(string expr, string domain, object criteria)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DCount", expr, domain, criteria);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx </remarks>
        /// <param name="expr">string expr</param>
        /// <param name="domain">string domain</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DCount(string expr, string domain)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DCount", expr, domain);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834705.aspx </remarks>
        /// <param name="stringExpr">string stringExpr</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Eval(string stringExpr)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Eval", stringExpr);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845778.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static string CurrentUser()
        {
            return Factory.ExecuteStringMethodGet(_instance, "CurrentUser");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196189.aspx </remarks>
        /// <param name="application">string application</param>
        /// <param name="topic">string topic</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object DDEInitiate(string application, string topic)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "DDEInitiate", application, topic);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197936.aspx </remarks>
        /// <param name="chanNum">object chanNum</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDEExecute(object chanNum, string command)
        {
            Factory.ExecuteMethod(_instance, "DDEExecute", chanNum, command);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194752.aspx </remarks>
        /// <param name="chanNum">object chanNum</param>
        /// <param name="item">string item</param>
        /// <param name="data">string data</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDEPoke(object chanNum, string item, string data)
        {
            Factory.ExecuteMethod(_instance, "DDEPoke", chanNum, item, data);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823145.aspx </remarks>
        /// <param name="chanNum">object chanNum</param>
        /// <param name="item">string item</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static string DDERequest(object chanNum, string item)
        {
            return Factory.ExecuteStringMethodGet(_instance, "DDERequest", chanNum, item);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197795.aspx </remarks>
        /// <param name="chanNum">object chanNum</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDETerminate(object chanNum)
        {
            Factory.ExecuteMethod(_instance, "DDETerminate", chanNum);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845193.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDETerminateAll()
        {
            Factory.ExecuteMethod(_instance, "DDETerminateAll");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835631.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.DAOApi.Database CurrentDb()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(_instance, "CurrentDb", NetOffice.DAOApi.Database.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196457.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.DAOApi.Database CodeDb()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(_instance, "CodeDb", NetOffice.DAOApi.Database.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="hwnd">Int32 hwnd</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void BeginUndoable(Int32 hwnd)
        {
            Factory.ExecuteMethod(_instance, "BeginUndoable", hwnd);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="yesno">Int16 yesno</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void SetUndoRecording(Int16 yesno)
        {
            Factory.ExecuteMethod(_instance, "SetUndoRecording", yesno);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845070.aspx </remarks>
        /// <param name="field">string field</param>
        /// <param name="fieldType">Int16 fieldType</param>
        /// <param name="expression">string expression</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static string BuildCriteria(string field, Int16 fieldType, string expression)
        {
            return Factory.ExecuteStringMethodGet(_instance, "BuildCriteria", field, fieldType, expression);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="moduleName">string moduleName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void InsertText(string text, string moduleName)
        {
            Factory.ExecuteMethod(_instance, "InsertText", text, moduleName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void ReloadAddIns()
        {
            Factory.ExecuteMethod(_instance, "ReloadAddIns");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836901.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.DAOApi.Workspace DefaultWorkspaceClone()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(_instance, "DefaultWorkspaceClone", NetOffice.DAOApi.Workspace.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197957.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void RefreshTitleBar()
        {
            Factory.ExecuteMethod(_instance, "RefreshTitleBar");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="changeFrom">string changeFrom</param>
        /// <param name="changeTo">string changeTo</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void AddAutoCorrect(string changeFrom, string changeTo)
        {
            Factory.ExecuteMethod(_instance, "AddAutoCorrect", changeFrom, changeTo);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="changeFrom">string changeFrom</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void DelAutoCorrect(string changeFrom)
        {
            Factory.ExecuteMethod(_instance, "DelAutoCorrect", changeFrom);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196179.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 hWndAccessApp()
        {
            return Factory.ExecuteInt32MethodGet(_instance, "hWndAccessApp");
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
        /// <param name="procedure">string procedure</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", procedure);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
        /// <param name="procedure">string procedure</param>
        /// <param name="arg1">optional object arg1</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", procedure, arg1);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
        /// <param name="procedure">string procedure</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", procedure, arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
        /// <param name="procedure">string procedure</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", procedure, arg1, arg2, arg3);
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx </remarks>
        /// <param name="value">object value</param>
        /// <param name="valueIfNull">optional object valueIfNull</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Nz(object value, object valueIfNull)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Nz", value, valueIfNull);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx </remarks>
        /// <param name="value">object value</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object Nz(object value)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Nz", value);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835072.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object LoadPicture(string fileName)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "LoadPicture", fileName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objtyp">Int32 objtyp</param>
        /// <param name="moduleName">string moduleName</param>
        /// <param name="fileName">string fileName</param>
        /// <param name="token">Int32 token</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void ReplaceModule(Int32 objtyp, string moduleName, string fileName, Int32 token)
        {
            Factory.ExecuteMethod(_instance, "ReplaceModule", objtyp, moduleName, fileName, token);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196488.aspx </remarks>
        /// <param name="errorNumber">object errorNumber</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object AccessError(object errorNumber)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "AccessError", errorNumber);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object BuilderString()
        {
            return Factory.ExecuteVariantMethodGet(_instance, "BuilderString");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193935.aspx </remarks>
        /// <param name="guid">object guid</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object StringFromGUID(object guid)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "StringFromGUID", guid);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197675.aspx </remarks>
        /// <param name="_string">object string</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object GUIDFromString(object _string)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "GUIDFromString", _string);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="id">Int32 id</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static object AppLoadString(Int32 id)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "AppLoadString", id);
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", new object[] { address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
        /// <param name="address">string address</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", address);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional string SubAddress = </param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address, object subAddress)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", address, subAddress);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional string SubAddress = </param>
        /// <param name="newWindow">optional bool NewWindow = false</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address, object subAddress, object newWindow)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", address, subAddress, newWindow);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional string SubAddress = </param>
        /// <param name="newWindow">optional bool NewWindow = false</param>
        /// <param name="addHistory">optional bool AddHistory = true</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", address, subAddress, newWindow, addHistory);
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", new object[] { address, subAddress, newWindow, addHistory, extraInfo });
        }

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
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
        {
            Factory.ExecuteMethod(_instance, "FollowHyperlink", new object[] { address, subAddress, newWindow, addHistory, extraInfo, method });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void SaveAsText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
        {
            Factory.ExecuteMethod(_instance, "SaveAsText", objectType, objectName, fileName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void LoadFromText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
        {
            Factory.ExecuteMethod(_instance, "LoadFromText", objectType, objectName, fileName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823011.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void AddToFavorites()
        {
            Factory.ExecuteMethod(_instance, "AddToFavorites");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194960.aspx </remarks>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void RefreshDatabaseWindow()
        {
            Factory.ExecuteMethod(_instance, "RefreshDatabaseWindow");
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191909.aspx </remarks>
        /// <param name="command">NetOffice.AccessApi.Enums.AcCommand command</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void RunCommand(NetOffice.AccessApi.Enums.AcCommand command)
        {
            Factory.ExecuteMethod(_instance, "RunCommand", command);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx </remarks>
        /// <param name="hyperlink">object hyperlink</param>
        /// <param name="part">optional NetOffice.AccessApi.Enums.AcHyperlinkPart Part = 0</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static string HyperlinkPart(object hyperlink, object part)
        {
            return Factory.ExecuteStringMethodGet(_instance, "HyperlinkPart", hyperlink, part);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx </remarks>
        /// <param name="hyperlink">object hyperlink</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static string HyperlinkPart(object hyperlink)
        {
            return Factory.ExecuteStringMethodGet(_instance, "HyperlinkPart", hyperlink);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821756.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static bool GetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "GetHiddenAttribute", objectType, objectName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822459.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        /// <param name="fHidden">bool fHidden</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void SetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, bool fHidden)
        {
            Factory.ExecuteMethod(_instance, "SetHiddenAttribute", objectType, objectName, fHidden);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="createNewFile">optional bool CreateNewFile = true</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName, object createNewFile)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.DataAccessPage>(_instance, "CreateDataAccessPage", NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType, fileName, createNewFile);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.DataAccessPage CreateDataAccessPage()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.DataAccessPage>(_instance, "CreateDataAccessPage", NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">optional object fileName</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.DataAccessPage>(_instance, "CreateDataAccessPage", NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType, fileName);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="connect">optional object connect</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void NewAccessProject(string filepath, object connect)
        {
            Factory.ExecuteMethod(_instance, "NewAccessProject", filepath, connect);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void NewAccessProject(string filepath)
        {
            Factory.ExecuteMethod(_instance, "NewAccessProject", filepath);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="exclusive">optional bool Exclusive = false</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void OpenAccessProject(string filepath, object exclusive)
        {
            Factory.ExecuteMethod(_instance, "OpenAccessProject", filepath, exclusive);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void OpenAccessProject(string filepath)
        {
            Factory.ExecuteMethod(_instance, "OpenAccessProject", filepath);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        /// <param name="connect">optional object connect</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void CreateAccessProject(string filepath, object connect)
        {
            Factory.ExecuteMethod(_instance, "CreateAccessProject", filepath, connect);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx </remarks>
        /// <param name="filepath">string filepath</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static void CreateAccessProject(string filepath)
        {
            Factory.ExecuteMethod(_instance, "CreateAccessProject", filepath);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        /// <param name="fullPrecision">optional object fullPrecision</param>
        /// <param name="triangulationPrecision">optional object triangulationPrecision</param>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
        {
            return Factory.ExecuteDoubleMethodGet(_instance, "EuroConvert", new object[] { number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision });
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
        {
            return Factory.ExecuteDoubleMethodGet(_instance, "EuroConvert", number, sourceCurrency, targetCurrency);
        }

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        /// <param name="fullPrecision">optional object fullPrecision</param>
        [CustomMethod]
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public static Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
        {
            return Factory.ExecuteDoubleMethodGet(_instance, "EuroConvert", number, sourceCurrency, targetCurrency, fullPrecision);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filepath">string filepath</param>
        /// <param name="exclusive">optional bool Exclusive = false</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void OpenCurrentDatabaseOld(string filepath, object exclusive)
        {
            Factory.ExecuteMethod(_instance, "OpenCurrentDatabaseOld", filepath, exclusive);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filepath">string filepath</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void OpenCurrentDatabaseOld(string filepath)
        {
            Factory.ExecuteMethod(_instance, "OpenCurrentDatabaseOld", filepath);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="path">optional string Path =  </param>
        /// <param name="name">optional string Name =  </param>
        /// <param name="company">optional string Company =  </param>
        /// <param name="workgroupID">optional string WorkgroupID =  </param>
        /// <param name="replace">optional bool Replace = false</param>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID, object replace)
        {
            Factory.ExecuteMethod(_instance, "CreateNewWorkgroupFile", new object[] { path, name, company, workgroupID, replace });
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void CreateNewWorkgroupFile()
        {
            Factory.ExecuteMethod(_instance, "CreateNewWorkgroupFile");
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="path">optional string Path =  </param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void CreateNewWorkgroupFile(object path)
        {
            Factory.ExecuteMethod(_instance, "CreateNewWorkgroupFile", path);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="path">optional string Path =  </param>
        /// <param name="name">optional string Name =  </param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void CreateNewWorkgroupFile(object path, object name)
        {
            Factory.ExecuteMethod(_instance, "CreateNewWorkgroupFile", path, name);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="path">optional string Path =  </param>
        /// <param name="name">optional string Name =  </param>
        /// <param name="company">optional string Company =  </param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void CreateNewWorkgroupFile(object path, object name, object company)
        {
            Factory.ExecuteMethod(_instance, "CreateNewWorkgroupFile", path, name, company);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="path">optional string Path =  </param>
        /// <param name="name">optional string Name =  </param>
        /// <param name="company">optional string Company =  </param>
        /// <param name="workgroupID">optional string WorkgroupID =  </param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID)
        {
            Factory.ExecuteMethod(_instance, "CreateNewWorkgroupFile", path, name, company, workgroupID);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195103.aspx </remarks>
        /// <param name="path">string path</param>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void SetDefaultWorkgroupFile(string path)
        {
            Factory.ExecuteMethod(_instance, "SetDefaultWorkgroupFile", path);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193465.aspx </remarks>
        /// <param name="sourceFilename">string sourceFilename</param>
        /// <param name="destinationFilename">string destinationFilename</param>
        /// <param name="destinationFileFormat">NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat</param>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ConvertAccessProject(string sourceFilename, string destinationFilename, NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat)
        {
            Factory.ExecuteMethod(_instance, "ConvertAccessProject", sourceFilename, destinationFilename, destinationFileFormat);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx </remarks>
        /// <param name="sourceFile">string sourceFile</param>
        /// <param name="destinationFile">string destinationFile</param>
        /// <param name="logFile">optional bool LogFile = false</param>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static bool CompactRepair(string sourceFile, string destinationFile, object logFile)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CompactRepair", sourceFile, destinationFile, logFile);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx </remarks>
        /// <param name="sourceFile">string sourceFile</param>
        /// <param name="destinationFile">string destinationFile</param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static bool CompactRepair(string sourceFile, string destinationFile)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CompactRepair", sourceFile, destinationFile);
        }

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
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags });
        }

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
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition, object additionalData)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition, additionalData });
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
        /// <param name="dataSource">string dataSource</param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", objectType, dataSource);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="dataTarget">optional string DataTarget = </param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", objectType, dataSource, dataTarget);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="dataTarget">optional string DataTarget = </param>
        /// <param name="schemaTarget">optional string SchemaTarget = </param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", objectType, dataSource, dataTarget, schemaTarget);
        }

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
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget });
        }

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
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget });
        }

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
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding });
        }

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
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition)
        {
            Factory.ExecuteMethod(_instance, "ExportXML", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition });
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx </remarks>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="importOptions">optional NetOffice.AccessApi.Enums.AcImportXMLOption ImportOptions = 1</param>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ImportXML(string dataSource, object importOptions)
        {
            Factory.ExecuteMethod(_instance, "ImportXML", dataSource, importOptions);
        }

        /// <summary>
        /// SupportByVersion Access 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx </remarks>
        /// <param name="dataSource">string dataSource</param>
        [CustomMethod]
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public static void ImportXML(string dataSource)
        {
            Factory.ExecuteMethod(_instance, "ImportXML", dataSource);
        }

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
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags });
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
        /// <param name="dataSource">string dataSource</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", objectType, dataSource);
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="dataTarget">optional string DataTarget = </param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", objectType, dataSource, dataTarget);
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="dataTarget">optional string DataTarget = </param>
        /// <param name="schemaTarget">optional string SchemaTarget = </param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", objectType, dataSource, dataTarget, schemaTarget);
        }

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
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget });
        }

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
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget });
        }

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
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
        {
            Factory.ExecuteMethod(_instance, "ExportXMLOld", new object[] { objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding });
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="transformSource">string transformSource</param>
        /// <param name="outputTarget">string outputTarget</param>
        /// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
        /// <param name="scriptOption">optional NetOffice.AccessApi.Enums.AcTransformXMLScriptOption ScriptOption = 1</param>
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput, object scriptOption)
        {
            Factory.ExecuteMethod(_instance, "TransformXML", new object[] { dataSource, transformSource, outputTarget, wellFormedXMLOutput, scriptOption });
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="transformSource">string transformSource</param>
        /// <param name="outputTarget">string outputTarget</param>
        [CustomMethod]
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void TransformXML(string dataSource, string transformSource, string outputTarget)
        {
            Factory.ExecuteMethod(_instance, "TransformXML", dataSource, transformSource, outputTarget);
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
        /// <param name="dataSource">string dataSource</param>
        /// <param name="transformSource">string transformSource</param>
        /// <param name="outputTarget">string outputTarget</param>
        /// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
        [CustomMethod]
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput)
        {
            Factory.ExecuteMethod(_instance, "TransformXML", dataSource, transformSource, outputTarget, wellFormedXMLOutput);
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834773.aspx </remarks>
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.AccessApi._AdditionalData CreateAdditionalData()
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.AccessApi._AdditionalData>(_instance, "CreateAdditionalData");
        }

        /// <summary>
        /// SupportByVersion Access 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dispid">Int32 dispid</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 11, 12, 14, 15, 16)]
        public static bool IsMemberSafe(Int32 dispid)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "IsMemberSafe", dispid);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="filepath">string filepath</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void NewCurrentDatabaseOld(string filepath)
        {
            Factory.ExecuteMethod(_instance, "NewCurrentDatabaseOld", filepath);
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left, top, width, height });
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="formName">string formName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, formName, controlType);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="formName">string formName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, formName, controlType, section);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="formName">string formName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        /// <param name="parent">optional object parent</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, formName, controlType, section, parent);
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left, top });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, columnName, left, top, width });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left, top, width, height });
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, reportName, controlType);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, reportName, controlType, section);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="reportName">string reportName</param>
        /// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
        /// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
        /// <param name="parent">optional object parent</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, reportName, controlType, section, parent);
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left, top });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, columnName, left, top, width });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateControlExOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateControlExOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { formName, controlType, section, parent, controlSource, left, top, width, height });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static NetOffice.AccessApi.Control CreateReportControlExOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(_instance, "CreateReportControlExOld", NetOffice.AccessApi.Control.LateBindingApiWrapperType, new object[] { reportName, controlType, section, parent, controlName, left, top, width, height });
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx </remarks>
        /// <param name="richText">object richText</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static string PlainText(object richText, object length)
        {
            return Factory.ExecuteStringMethodGet(_instance, "PlainText", richText, length);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx </remarks>
        /// <param name="richText">object richText</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static string PlainText(object richText)
        {
            return Factory.ExecuteStringMethodGet(_instance, "PlainText", richText);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx </remarks>
        /// <param name="plainText">object plainText</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static string HtmlEncode(object plainText, object length)
        {
            return Factory.ExecuteStringMethodGet(_instance, "HtmlEncode", plainText, length);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx </remarks>
        /// <param name="plainText">object plainText</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static string HtmlEncode(object plainText)
        {
            return Factory.ExecuteStringMethodGet(_instance, "HtmlEncode", plainText);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194416.aspx </remarks>
        /// <param name="customUIName">string customUIName</param>
        /// <param name="customUIXML">string customUIXML</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void LoadCustomUI(string customUIName, string customUIXML)
        {
            Factory.ExecuteMethod(_instance, "LoadCustomUI", customUIName, customUIXML);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193467.aspx </remarks>
        /// <param name="path">string path</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ExportNavigationPane(string path)
        {
            Factory.ExecuteMethod(_instance, "ExportNavigationPane", path);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx </remarks>
        /// <param name="path">string path</param>
        /// <param name="fAppendOnly">optional bool fAppendOnly = false</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ImportNavigationPane(string path, object fAppendOnly)
        {
            Factory.ExecuteMethod(_instance, "ImportNavigationPane", path, fAppendOnly);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx </remarks>
        /// <param name="path">string path</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ImportNavigationPane(string path)
        {
            Factory.ExecuteMethod(_instance, "ImportNavigationPane", path);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835727.aspx </remarks>
        /// <param name="tableName">string tableName</param>
        /// <param name="columnName">string columnName</param>
        /// <param name="queryString">string queryString</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static string ColumnHistory(string tableName, string columnName, string queryString)
        {
            return Factory.ExecuteStringMethodGet(_instance, "ColumnHistory", tableName, columnName, queryString);
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage, object toPage)
        {
            Factory.ExecuteMethod(_instance, "ExportCustomFixedFormat", new object[] { externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage, toPage });
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="externalExporter">object externalExporter</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="objectName">string objectName</param>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
        {
            Factory.ExecuteMethod(_instance, "ExportCustomFixedFormat", externalExporter, outputFileName, objectName, objectType);
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords)
        {
            Factory.ExecuteMethod(_instance, "ExportCustomFixedFormat", new object[] { externalExporter, outputFileName, objectName, objectType, selectedRecords });
        }

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
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage)
        {
            Factory.ExecuteMethod(_instance, "ExportCustomFixedFormat", new object[] { externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage });
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821429.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
        {
            Factory.ExecuteMethod(_instance, "SaveAsAXL", objectType, objectName, fileName);
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845765.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static void LoadFromAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
        {
            Factory.ExecuteMethod(_instance, "LoadFromAXL", objectType, objectName, fileName);
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData, object variation)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData, variation });
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category });
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category, previewPath });
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category, previewPath, description });
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category, previewPath, description, instantiationForm });
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart });
        }

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
        [SupportByVersion("Access", 14, 15, 16)]
        public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData)
        {
            Factory.ExecuteMethod(_instance, "SaveAsTemplate", new object[] { path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData });
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835421.aspx </remarks>
        /// <param name="path">string path</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static void InstantiateTemplate(string path)
        {
            Factory.ExecuteMethod(_instance, "InstantiateTemplate", path);
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834388.aspx </remarks>
        /// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static object CurrentWebUser(NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CurrentWebUser", displayOption);
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836539.aspx </remarks>
        /// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static object CurrentWebUserGroups(NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CurrentWebUserGroups", displayOption);
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193453.aspx </remarks>
        /// <param name="groupNameOrID">object groupNameOrID</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static bool IsCurrentWebUserInGroup(object groupNameOrID)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "IsCurrentWebUserInGroup", groupNameOrID);
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834368.aspx </remarks>
        /// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
        /// <param name="objectName">string objectName</param>
        [SupportByVersion("Access", 14, 15, 16)]
        public static void DirtyObject(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
        {
            Factory.ExecuteMethod(_instance, "DirtyObject", objectType, objectName);
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 14, 15, 16)]
        public static bool IsClient()
        {
            return Factory.ExecuteBoolMethodGet(_instance, "IsClient");
        }
        
        #endregion
    }
}
