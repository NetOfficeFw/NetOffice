using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface _Application 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Application : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Application(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document ActiveDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Document>(this, "ActiveDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Window ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Window>(this, "ActiveWindow", NetOffice.PublisherApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Build
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorSchemes ColorSchemes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorSchemes>(this, "ColorSchemes", NetOffice.PublisherApi.ColorSchemes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(this, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_FileDialog")]
		public NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
		{
			return get_FileDialog(type);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(this, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Language
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Language");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Options Options
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Options>(this, "Options", NetOffice.PublisherApi.Options.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string PathSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PathSeparator");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string ProductCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProductCode");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintPreview
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintPreview");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ScreenUpdating
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ScreenUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Selection Selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Selection>(this, "Selection", NetOffice.PublisherApi.Selection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool SnapToGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SnapToGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapToGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool SnapToObjects
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SnapToObjects");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapToObjects", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string TemplateFolderPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TemplateFolderPath");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.OfficeDataSourceObject OfficeDataSourceObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeDataSourceObject>(this, "OfficeDataSourceObject", NetOffice.OfficeApi.OfficeDataSourceObject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool WizardCatalogVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WizardCatalogVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WizardCatalogVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Documents Documents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Documents>(this, "Documents", NetOffice.PublisherApi.Documents.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebOptions WebOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebOptions>(this, "WebOptions", NetOffice.PublisherApi.WebOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.InstalledPrinters InstalledPrinters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.InstalledPrinters>(this, "InstalledPrinters", NetOffice.PublisherApi.InstalledPrinters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(this, "MsoDebugOptions", NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ValidateAddressVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ValidateAddressVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ValidateAddressVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool InsertBarcodeVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InsertBarcodeVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InsertBarcodeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string ShowFollowUpCustom
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShowFollowUpCustom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFollowUpCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutomationSecurity>(this, "AutomationSecurity");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutomationSecurity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CaptionStyles CaptionStyles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CaptionStyles>(this, "CaptionStyles", NetOffice.PublisherApi.CaptionStyles.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="dir">string dir</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ChangeFileOpenDirectory(string dir)
		{
			 Factory.ExecuteMethod(this, "ChangeFileOpenDirectory", dir);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="helpType">NetOffice.PublisherApi.Enums.PbHelpType helpType</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Help(NetOffice.PublisherApi.Enums.PbHelpType helpType)
		{
			 Factory.ExecuteMethod(this, "Help", helpType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="_object">object object</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsValidObject(object _object)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsValidObject", _object);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document NewDocument(object wizard, object design)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "NewDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, wizard, design);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document NewDocument()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "NewDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document NewDocument(object wizard)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "NewDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, wizard);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="readOnly">optional bool ReadOnly = false</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		/// <param name="saveChanges">optional NetOffice.PublisherApi.Enums.PbSaveOptions SaveChanges = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Open(string filename, object readOnly, object addToRecentFiles, object saveChanges)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename, readOnly, addToRecentFiles, saveChanges);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Open(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="readOnly">optional bool ReadOnly = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Open(string filename, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename, readOnly);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="readOnly">optional bool ReadOnly = false</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Open(string filename, object readOnly, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void LaunchWebService()
		{
			 Factory.ExecuteMethod(this, "LaunchWebService");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single CentimetersToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "CentimetersToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single EmusToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "EmusToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single InchesToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "InchesToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single LinesToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "LinesToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single MillimetersToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "MillimetersToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PicasToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PicasToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PixelsToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PixelsToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single TwipsToPoints(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "TwipsToPoints", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToCentimeters(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToCentimeters", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToEmus(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToEmus", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToInches(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToInches", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToLines(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToLines", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToMillimeters(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToMillimeters", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToPicas(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToPicas", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToPixels(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToPixels", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">Single value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single PointsToTwips(Single value)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToTwips", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizardCatalog(object wizard)
		{
			 Factory.ExecuteMethod(this, "ShowWizardCatalog", wizard);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizardCatalog()
		{
			 Factory.ExecuteMethod(this, "ShowWizardCatalog");
		}

		#endregion

		#pragma warning restore
	}
}
