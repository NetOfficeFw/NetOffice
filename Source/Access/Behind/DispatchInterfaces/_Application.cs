using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CoreServices;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _Application
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Application : COMObject, NetOffice.AccessApi._Application
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.AccessApi._Application);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_Application);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Application() : base()
		{

		}

		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192087.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", typeof(NetOffice.AccessApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836400.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822407.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object CodeContextObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CodeContextObject");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835352.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string MenuBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MenuBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845319.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CurrentObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196795.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string CurrentObjectName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrentObjectName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837183.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Forms Forms
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Forms>(this, "Forms", typeof(NetOffice.AccessApi.Forms));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834339.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Reports Reports
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Reports>(this, "Reports", typeof(NetOffice.AccessApi.Reports));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835056.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Screen Screen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Screen>(this, "Screen", typeof(NetOffice.AccessApi.Screen));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845564.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.DoCmd DoCmd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.DoCmd>(this, "DoCmd", typeof(NetOffice.AccessApi.DoCmd));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195236.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ShortcutMenuBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821493.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Visible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836033.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool UserControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UserControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821724.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.DAOApi.DBEngine DBEngine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.DBEngine>(this, "DBEngine", typeof(NetOffice.DAOApi.DBEngine));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821379.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", typeof(NetOffice.OfficeApi.Assistant));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835326.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.References References
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.References>(this, "References", typeof(NetOffice.AccessApi.References));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836265.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Modules Modules
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Modules>(this, "Modules", typeof(NetOffice.AccessApi.Modules));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(this, "FileSearch", typeof(NetOffice.OfficeApi.FileSearch));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823044.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool IsCompiled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsCompiled");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822476.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", typeof(NetOffice.VBIDEApi.VBE));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.DataAccessPages DataAccessPages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.DataAccessPages>(this, "DataAccessPages", typeof(NetOffice.AccessApi.DataAccessPages));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ADOConnectString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ADOConnectString");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193770.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.CurrentProject CurrentProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CurrentProject>(this, "CurrentProject", typeof(NetOffice.AccessApi.CurrentProject));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193230.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.CurrentData CurrentData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CurrentData>(this, "CurrentData", typeof(NetOffice.AccessApi.CurrentData));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197047.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.CodeProject CodeProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CodeProject>(this, "CodeProject", typeof(NetOffice.AccessApi.CodeProject));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836912.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.CodeData CodeData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.CodeData>(this, "CodeData", typeof(NetOffice.AccessApi.CodeData));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.WizHook WizHook
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.WizHook>(this, "WizHook", typeof(NetOffice.AccessApi.WizHook));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822077.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ProductCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProductCode");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822463.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", typeof(NetOffice.OfficeApi.COMAddIns));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194961.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.DefaultWebOptions DefaultWebOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.DefaultWebOptions>(this, "DefaultWebOptions", typeof(NetOffice.AccessApi.DefaultWebOptions));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836634.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(this, "LanguageSettings", typeof(NetOffice.OfficeApi.LanguageSettings));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(this, "AnswerWizard", typeof(NetOffice.OfficeApi.AnswerWizard));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822721.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFeatureInstall>(this, "FeatureInstall");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FeatureInstall", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object VGXFrameInterval
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "VGXFrameInterval");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx </remarks>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(this, "FileDialog", typeof(NetOffice.OfficeApi.FileDialog), dialogType);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx </remarks>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersion("Access", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		public virtual NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return get_FileDialog(dialogType);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845884.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool BrokenReference
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BrokenReference");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195779.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Printers Printers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Printers>(this, "Printers", typeof(NetOffice.AccessApi.Printers));
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821394.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.AccessApi._Printer Printer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Printer>(this, "Printer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Printer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(this, "MsoDebugOptions", typeof(NetOffice.OfficeApi.MsoDebugOptions));
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192859.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835096.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 Build
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191715.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.NewFile NewFileTaskPane
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(this, "NewFileTaskPane", typeof(NetOffice.OfficeApi.NewFile));
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845345.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.AccessApi._AutoCorrect AutoCorrect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._AutoCorrect>(this, "AutoCorrect");
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193178.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutomationSecurity>(this, "AutomationSecurity");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AutomationSecurity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845034.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.MacroError MacroError
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.MacroError>(this, "MacroError", typeof(NetOffice.AccessApi.MacroError));
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192459.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.TempVars TempVars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.TempVars>(this, "TempVars", typeof(NetOffice.AccessApi.TempVars));
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192450.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", typeof(NetOffice.OfficeApi.IAssistance));
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837286.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public virtual NetOffice.AccessApi.WebServices WebServices
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.WebServices>(this, "WebServices", typeof(NetOffice.AccessApi.WebServices));
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.LocalVars LocalVars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.LocalVars>(this, "LocalVars", typeof(NetOffice.AccessApi.LocalVars));
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj249062.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public virtual NetOffice.AccessApi.ReturnVars ReturnVars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.ReturnVars>(this, "ReturnVars", typeof(NetOffice.AccessApi.ReturnVars));
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
		public virtual void NewCurrentDatabase(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewCurrentDatabase", filepath);
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
		public virtual void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress, object listID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewCurrentDatabase", new object[]{ filepath, fileFormat, template, siteAddress, listID });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void NewCurrentDatabase(string filepath, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewCurrentDatabase", filepath, fileFormat);
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
		public virtual void NewCurrentDatabase(string filepath, object fileFormat, object template)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewCurrentDatabase", filepath, fileFormat, template);
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
		public virtual void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewCurrentDatabase", filepath, fileFormat, template, siteAddress);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenCurrentDatabase(string filepath, object exclusive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenCurrentDatabase", filepath, exclusive);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		/// <param name="bstrPassword">optional string bstrPassword = </param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenCurrentDatabase(string filepath, object exclusive, object bstrPassword)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenCurrentDatabase", filepath, exclusive, bstrPassword);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenCurrentDatabase(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenCurrentDatabase", filepath);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192308.aspx </remarks>
		/// <param name="optionName">string optionName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object GetOption(string optionName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetOption", optionName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195513.aspx </remarks>
		/// <param name="optionName">string optionName</param>
		/// <param name="setting">object setting</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetOption(string optionName, object setting)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOption", optionName, setting);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx </remarks>
		/// <param name="echoOn">Int16 echoOn</param>
		/// <param name="bstrStatusBarText">optional string bstrStatusBarText = </param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Echo(Int16 echoOn, object bstrStatusBarText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Echo", echoOn, bstrStatusBarText);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx </remarks>
		/// <param name="echoOn">Int16 echoOn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Echo(Int16 echoOn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Echo", echoOn);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836850.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CloseCurrentDatabase()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CloseCurrentDatabase");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx </remarks>
		/// <param name="option">optional NetOffice.AccessApi.Enums.AcQuitOption Option = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Quit(object option)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit", option);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Quit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		/// <param name="argument2">optional object argument2</param>
		/// <param name="argument3">optional object argument3</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2, object argument3)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SysCmd", action, argument2, argument3);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SysCmd", action);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx </remarks>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		/// <param name="argument2">optional object argument2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SysCmd", action, argument2);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
		/// <param name="database">optional object database</param>
		/// <param name="formTemplate">optional object formTemplate</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Form CreateForm(object database, object formTemplate)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Form>(this, "CreateForm", typeof(NetOffice.AccessApi.Form), database, formTemplate);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Form CreateForm()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Form>(this, "CreateForm", typeof(NetOffice.AccessApi.Form));
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx </remarks>
		/// <param name="database">optional object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Form CreateForm(object database)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Form>(this, "CreateForm", typeof(NetOffice.AccessApi.Form), database);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
		/// <param name="database">optional object database</param>
		/// <param name="reportTemplate">optional object reportTemplate</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Report CreateReport(object database, object reportTemplate)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Report>(this, "CreateReport", typeof(NetOffice.AccessApi.Report), database, reportTemplate);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Report CreateReport()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Report>(this, "CreateReport", typeof(NetOffice.AccessApi.Report));
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx </remarks>
		/// <param name="database">optional object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Report CreateReport(object database)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Report>(this, "CreateReport", typeof(NetOffice.AccessApi.Report), database);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), formName, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), formName, controlType, section);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), formName, controlType, section, parent);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left, top });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControl", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left, top, width });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), reportName, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), reportName, controlType, section);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), reportName, controlType, section, parent);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left, top });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControl", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left, top, width });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlEx(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlEx", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, controlSource, left, top, width, height });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlEx(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlEx", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, controlName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836733.aspx </remarks>
		/// <param name="formName">string formName</param>
		/// <param name="controlName">string controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DeleteControl(string formName, string controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteControl", formName, controlName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191904.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlName">string controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DeleteReportControl(string reportName, string controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteReportControl", reportName, controlName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197044.aspx </remarks>
		/// <param name="reportName">string reportName</param>
		/// <param name="expression">string expression</param>
		/// <param name="header">Int16 header</param>
		/// <param name="footer">Int16 footer</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CreateGroupLevel(string reportName, string expression, Int16 header, Int16 footer)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CreateGroupLevel", reportName, expression, header, footer);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DMin(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DMin", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DMin(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DMin", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DMax(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DMax", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DMax(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DMax", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DSum(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DSum", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DSum(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DSum", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DAvg(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DAvg", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DAvg(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DAvg", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DLookup(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DLookup", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DLookup(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DLookup", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DLast(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DLast", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DLast(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DLast", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DVar(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DVar", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DVar(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DVar", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DVarP(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DVarP", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DVarP(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DVarP", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DStDev(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DStDev", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DStDev(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DStDev", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DStDevP(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DStDevP", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DStDevP(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DStDevP", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DFirst(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DFirst", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DFirst(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DFirst", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DCount(string expr, string domain, object criteria)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DCount", expr, domain, criteria);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx </remarks>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DCount(string expr, string domain)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DCount", expr, domain);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834705.aspx </remarks>
		/// <param name="stringExpr">string stringExpr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Eval(string stringExpr)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Eval", stringExpr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845778.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string CurrentUser()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CurrentUser");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196189.aspx </remarks>
		/// <param name="application">string application</param>
		/// <param name="topic">string topic</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object DDEInitiate(string application, string topic)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DDEInitiate", application, topic);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197936.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="command">string command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DDEExecute(object chanNum, string command)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DDEExecute", chanNum, command);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194752.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="item">string item</param>
		/// <param name="data">string data</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DDEPoke(object chanNum, string item, string data)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DDEPoke", chanNum, item, data);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823145.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="item">string item</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string DDERequest(object chanNum, string item)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "DDERequest", chanNum, item);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197795.aspx </remarks>
		/// <param name="chanNum">object chanNum</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DDETerminate(object chanNum)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DDETerminate", chanNum);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845193.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DDETerminateAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DDETerminateAll");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835631.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.DAOApi.Database CurrentDb()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CurrentDb", typeof(NetOffice.DAOApi.Database));
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196457.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.DAOApi.Database CodeDb()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CodeDb", typeof(NetOffice.DAOApi.Database));
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwnd">Int32 hwnd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void BeginUndoable(Int32 hwnd)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginUndoable", hwnd);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="yesno">Int16 yesno</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetUndoRecording(Int16 yesno)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetUndoRecording", yesno);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845070.aspx </remarks>
		/// <param name="field">string field</param>
		/// <param name="fieldType">Int16 fieldType</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string BuildCriteria(string field, Int16 fieldType, string expression)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "BuildCriteria", field, fieldType, expression);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="moduleName">string moduleName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void InsertText(string text, string moduleName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertText", text, moduleName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ReloadAddIns()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReloadAddIns");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836901.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.DAOApi.Workspace DefaultWorkspaceClone()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "DefaultWorkspaceClone", typeof(NetOffice.DAOApi.Workspace));
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197957.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RefreshTitleBar()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshTitleBar");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="changeFrom">string changeFrom</param>
		/// <param name="changeTo">string changeTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void AddAutoCorrect(string changeFrom, string changeTo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddAutoCorrect", changeFrom, changeTo);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="changeFrom">string changeFrom</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DelAutoCorrect(string changeFrom)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DelAutoCorrect", changeFrom);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196179.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 hWndAccessApp()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "hWndAccessApp");
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", procedure);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", procedure, arg1);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx </remarks>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", procedure, arg1, arg2);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", procedure, arg1, arg2, arg3);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[]{ procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx </remarks>
		/// <param name="value">object value</param>
		/// <param name="valueIfNull">optional object valueIfNull</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Nz(object value, object valueIfNull)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Nz", value, valueIfNull);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx </remarks>
		/// <param name="value">object value</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object Nz(object value)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Nz", value);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835072.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object LoadPicture(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LoadPicture", fileName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objtyp">Int32 objtyp</param>
		/// <param name="moduleName">string moduleName</param>
		/// <param name="fileName">string fileName</param>
		/// <param name="token">Int32 token</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ReplaceModule(Int32 objtyp, string moduleName, string fileName, Int32 token)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceModule", objtyp, moduleName, fileName, token);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196488.aspx </remarks>
		/// <param name="errorNumber">object errorNumber</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object AccessError(object errorNumber)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AccessError", errorNumber);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object BuilderString()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BuilderString");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193935.aspx </remarks>
		/// <param name="guid">object guid</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object StringFromGUID(object guid)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "StringFromGUID", guid);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197675.aspx </remarks>
		/// <param name="_string">object string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object GUIDFromString(object _string)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GUIDFromString", _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="id">Int32 id</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object AppLoadString(Int32 id)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AppLoadString", id);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address, object subAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo });
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SaveAsText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsText", objectType, objectName, fileName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void LoadFromText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LoadFromText", objectType, objectName, fileName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823011.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void AddToFavorites()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194960.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RefreshDatabaseWindow()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshDatabaseWindow");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191909.aspx </remarks>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunCommand(NetOffice.AccessApi.Enums.AcCommand command)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunCommand", command);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx </remarks>
		/// <param name="hyperlink">object hyperlink</param>
		/// <param name="part">optional NetOffice.AccessApi.Enums.AcHyperlinkPart Part = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string HyperlinkPart(object hyperlink, object part)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "HyperlinkPart", hyperlink, part);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx </remarks>
		/// <param name="hyperlink">object hyperlink</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string HyperlinkPart(object hyperlink)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "HyperlinkPart", hyperlink);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821756.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool GetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetHiddenAttribute", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822459.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fHidden">bool fHidden</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, bool fHidden)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetHiddenAttribute", objectType, objectName, fHidden);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="createNewFile">optional bool CreateNewFile = true</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName, object createNewFile)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.DataAccessPage>(this, "CreateDataAccessPage", typeof(NetOffice.AccessApi.DataAccessPage), fileName, createNewFile);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.DataAccessPage CreateDataAccessPage()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.DataAccessPage>(this, "CreateDataAccessPage", typeof(NetOffice.AccessApi.DataAccessPage));
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.DataAccessPage>(this, "CreateDataAccessPage", typeof(NetOffice.AccessApi.DataAccessPage), fileName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void NewAccessProject(string filepath, object connect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewAccessProject", filepath, connect);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void NewAccessProject(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewAccessProject", filepath);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenAccessProject(string filepath, object exclusive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenAccessProject", filepath, exclusive);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenAccessProject(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenAccessProject", filepath);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CreateAccessProject(string filepath, object connect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateAccessProject", filepath, connect);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx </remarks>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CreateAccessProject(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateAccessProject", filepath);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
		{
			return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EuroConvert", new object[]{ number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx </remarks>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
		{
			return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EuroConvert", number, sourceCurrency, targetCurrency);
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
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
		{
			return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EuroConvert", number, sourceCurrency, targetCurrency, fullPrecision);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenCurrentDatabaseOld(string filepath, object exclusive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenCurrentDatabaseOld", filepath, exclusive);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenCurrentDatabaseOld(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenCurrentDatabaseOld", filepath);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		/// <param name="replace">optional bool Replace = false</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID, object replace)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateNewWorkgroupFile", new object[]{ path, name, company, workgroupID, replace });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CreateNewWorkgroupFile()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateNewWorkgroupFile");
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CreateNewWorkgroupFile(object path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateNewWorkgroupFile", path);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CreateNewWorkgroupFile(object path, object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateNewWorkgroupFile", path, name);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CreateNewWorkgroupFile(object path, object name, object company)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateNewWorkgroupFile", path, name, company);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateNewWorkgroupFile", path, name, company, workgroupID);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195103.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void SetDefaultWorkgroupFile(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultWorkgroupFile", path);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193465.aspx </remarks>
		/// <param name="sourceFilename">string sourceFilename</param>
		/// <param name="destinationFilename">string destinationFilename</param>
		/// <param name="destinationFileFormat">NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ConvertAccessProject(string sourceFilename, string destinationFilename, NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertAccessProject", sourceFilename, destinationFilename, destinationFileFormat);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx </remarks>
		/// <param name="sourceFile">string sourceFile</param>
		/// <param name="destinationFile">string destinationFile</param>
		/// <param name="logFile">optional bool LogFile = false</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool CompactRepair(string sourceFile, string destinationFile, object logFile)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CompactRepair", sourceFile, destinationFile, logFile);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx </remarks>
		/// <param name="sourceFile">string sourceFile</param>
		/// <param name="destinationFile">string destinationFile</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool CompactRepair(string sourceFile, string destinationFile)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CompactRepair", sourceFile, destinationFile);
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
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags });
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition, object additionalData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition, additionalData });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", objectType, dataSource);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", objectType, dataSource, dataTarget);
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
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", objectType, dataSource, dataTarget, schemaTarget);
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
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget });
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
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget });
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
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding });
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="importOptions">optional NetOffice.AccessApi.Enums.AcImportXMLOption ImportOptions = 1</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ImportXML(string dataSource, object importOptions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ImportXML", dataSource, importOptions);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void ImportXML(string dataSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ImportXML", dataSource);
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags });
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", objectType, dataSource);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", objectType, dataSource, dataTarget);
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", objectType, dataSource, dataTarget, schemaTarget);
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget });
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget });
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXMLOld", new object[]{ objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding });
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput, object scriptOption)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransformXML", new object[]{ dataSource, transformSource, outputTarget, wellFormedXMLOutput, scriptOption });
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx </remarks>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void TransformXML(string dataSource, string transformSource, string outputTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransformXML", dataSource, transformSource, outputTarget);
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
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransformXML", dataSource, transformSource, outputTarget, wellFormedXMLOutput);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834773.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.AccessApi._AdditionalData CreateAdditionalData()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.AccessApi._AdditionalData>(this, "CreateAdditionalData");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual bool IsMemberSafe(Int32 dispid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void NewCurrentDatabaseOld(string filepath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewCurrentDatabaseOld", filepath);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), formName, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), formName, controlType, section);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), formName, controlType, section, parent);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left, top });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, columnName, left, top, width });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), reportName, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), reportName, controlType, section);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), reportName, controlType, section, parent);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left, top });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlOld", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, columnName, left, top, width });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateControlExOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateControlExOld", typeof(NetOffice.AccessApi.Control), new object[]{ formName, controlType, section, parent, controlSource, left, top, width, height });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Control CreateReportControlExOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.AccessApi.Control>(this, "CreateReportControlExOld", typeof(NetOffice.AccessApi.Control), new object[]{ reportName, controlType, section, parent, controlName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx </remarks>
		/// <param name="richText">object richText</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string PlainText(object richText, object length)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "PlainText", richText, length);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx </remarks>
		/// <param name="richText">object richText</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string PlainText(object richText)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "PlainText", richText);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx </remarks>
		/// <param name="plainText">object plainText</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string HtmlEncode(object plainText, object length)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "HtmlEncode", plainText, length);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx </remarks>
		/// <param name="plainText">object plainText</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string HtmlEncode(object plainText)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "HtmlEncode", plainText);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194416.aspx </remarks>
		/// <param name="customUIName">string customUIName</param>
		/// <param name="customUIXML">string customUIXML</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void LoadCustomUI(string customUIName, string customUIXML)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LoadCustomUI", customUIName, customUIXML);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193467.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ExportNavigationPane(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportNavigationPane", path);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fAppendOnly">optional bool fAppendOnly = false</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ImportNavigationPane(string path, object fAppendOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ImportNavigationPane", path, fAppendOnly);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ImportNavigationPane(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ImportNavigationPane", path);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835727.aspx </remarks>
		/// <param name="tableName">string tableName</param>
		/// <param name="columnName">string columnName</param>
		/// <param name="queryString">string queryString</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string ColumnHistory(string tableName, string columnName, string queryString)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ColumnHistory", tableName, columnName, queryString);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage, object toPage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportCustomFixedFormat", new object[]{ externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage, toPage });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportCustomFixedFormat", externalExporter, outputFileName, objectName, objectType);
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportCustomFixedFormat", new object[]{ externalExporter, outputFileName, objectName, objectType, selectedRecords });
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
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportCustomFixedFormat", new object[]{ externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage });
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821429.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsAXL", objectType, objectName, fileName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845765.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void LoadFromAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LoadFromAXL", objectType, objectName, fileName);
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData, object variation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData, variation });
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category });
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category, previewPath });
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category, previewPath, description });
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category, previewPath, description, instantiationForm });
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart });
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
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsTemplate", new object[]{ path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData });
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835421.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void InstantiateTemplate(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InstantiateTemplate", path);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834388.aspx </remarks>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual object CurrentWebUser(NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CurrentWebUser", displayOption);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836539.aspx </remarks>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual object CurrentWebUserGroups(NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CurrentWebUserGroups", displayOption);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193453.aspx </remarks>
		/// <param name="groupNameOrID">object groupNameOrID</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual bool IsCurrentWebUserInGroup(object groupNameOrID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsCurrentWebUserInGroup", groupNameOrID);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834368.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void DirtyObject(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DirtyObject", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public virtual bool IsClient()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsClient");
		}

        #endregion

        #pragma warning restore
    }
}
