using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface _Application
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Application : COMObject, NetOffice.PublisherApi._Application
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
                    _contractType = typeof(NetOffice.PublisherApi._Application);
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
				return LateBindingApiWrapperType;
			}
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document ActiveDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Document>(this, "ActiveDocument", typeof(NetOffice.PublisherApi.Document));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Window ActiveWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Window>(this, "ActiveWindow", typeof(NetOffice.PublisherApi.Window));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", typeof(NetOffice.OfficeApi.Assistant));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Build
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorSchemes ColorSchemes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorSchemes>(this, "ColorSchemes", typeof(NetOffice.PublisherApi.ColorSchemes));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", typeof(NetOffice.OfficeApi.COMAddIns));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(this, "FileDialog", typeof(NetOffice.OfficeApi.FileDialog), type);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_FileDialog")]
		public virtual NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
		{
			return get_FileDialog(type);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(this, "FileSearch", typeof(NetOffice.OfficeApi.FileSearch));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Language
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Language");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Options Options
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Options>(this, "Options", typeof(NetOffice.PublisherApi.Options));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string PathSeparator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PathSeparator");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string ProductCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProductCode");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool PrintPreview
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintPreview");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ScreenUpdating
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ScreenUpdating");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Selection Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Selection>(this, "Selection", typeof(NetOffice.PublisherApi.Selection));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool SnapToGuides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SnapToGuides");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapToGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool SnapToObjects
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SnapToObjects");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapToObjects", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string TemplateFolderPath
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TemplateFolderPath");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.OfficeDataSourceObject OfficeDataSourceObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeDataSourceObject>(this, "OfficeDataSourceObject", typeof(NetOffice.OfficeApi.OfficeDataSourceObject));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool WizardCatalogVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WizardCatalogVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WizardCatalogVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Documents Documents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Documents>(this, "Documents", typeof(NetOffice.PublisherApi.Documents));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.WebOptions WebOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebOptions>(this, "WebOptions", typeof(NetOffice.PublisherApi.WebOptions));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.InstalledPrinters InstalledPrinters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.InstalledPrinters>(this, "InstalledPrinters", typeof(NetOffice.PublisherApi.InstalledPrinters));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(this, "MsoDebugOptions", typeof(NetOffice.OfficeApi.MsoDebugOptions));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ValidateAddressVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ValidateAddressVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ValidateAddressVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool InsertBarcodeVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InsertBarcodeVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InsertBarcodeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string ShowFollowUpCustom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShowFollowUpCustom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowFollowUpCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", typeof(NetOffice.OfficeApi.IAssistance));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.CaptionStyles CaptionStyles
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CaptionStyles>(this, "CaptionStyles", typeof(NetOffice.PublisherApi.CaptionStyles));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="dir">string dir</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ChangeFileOpenDirectory(string dir)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeFileOpenDirectory", dir);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="helpType">NetOffice.PublisherApi.Enums.PbHelpType helpType</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Help(NetOffice.PublisherApi.Enums.PbHelpType helpType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Help", helpType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="_object">object object</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool IsValidObject(object _object)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsValidObject", _object);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document NewDocument(object wizard, object design)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "NewDocument", typeof(NetOffice.PublisherApi.Document), wizard, design);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document NewDocument()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "NewDocument", typeof(NetOffice.PublisherApi.Document));
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document NewDocument(object wizard)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "NewDocument", typeof(NetOffice.PublisherApi.Document), wizard);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="readOnly">optional bool ReadOnly = false</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		/// <param name="saveChanges">optional NetOffice.PublisherApi.Enums.PbSaveOptions SaveChanges = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Open(string filename, object readOnly, object addToRecentFiles, object saveChanges)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", typeof(NetOffice.PublisherApi.Document), filename, readOnly, addToRecentFiles, saveChanges);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Open(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", typeof(NetOffice.PublisherApi.Document), filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="readOnly">optional bool ReadOnly = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Open(string filename, object readOnly)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", typeof(NetOffice.PublisherApi.Document), filename, readOnly);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="readOnly">optional bool ReadOnly = false</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Open(string filename, object readOnly, object addToRecentFiles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", typeof(NetOffice.PublisherApi.Document), filename, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Quit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void LaunchWebService()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LaunchWebService");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single CentimetersToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "CentimetersToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single EmusToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "EmusToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single InchesToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "InchesToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single LinesToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "LinesToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single MillimetersToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "MillimetersToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PicasToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PicasToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PixelsToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PixelsToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single TwipsToPoints(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "TwipsToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToCentimeters(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToCentimeters", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToEmus(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToEmus", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToInches(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToInches", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToLines(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToLines", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToMillimeters(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToMillimeters", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToPicas(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToPicas", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToPixels(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToPixels", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single PointsToTwips(Single value)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToTwips", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardCatalog(object wizard)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardCatalog", wizard);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardCatalog()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardCatalog");
		}

		#endregion

		#pragma warning restore
	}
}
