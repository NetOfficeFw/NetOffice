using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Publisher, 14,15,16
    ///</summary>
    [SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(PublisherApi.Application))]
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
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document ActiveDocument
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Document>(_instance, "ActiveDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Window ActiveWindow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Window>(_instance, "ActiveWindow", NetOffice.PublisherApi.Window.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(_instance, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.Assistant Assistant
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(_instance, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Int32 Build
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Build");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.ColorSchemes ColorSchemes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorSchemes>(_instance, "ColorSchemes", NetOffice.PublisherApi.ColorSchemes.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.OfficeApi.COMAddIns COMAddIns
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(_instance, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.OfficeApi.CommandBars CommandBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(_instance, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(_instance, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, type);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_FileDialog
        /// </summary>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_FileDialog")]
        public static NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
        {
            return get_FileDialog(type);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.FileSearch FileSearch
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(_instance, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Int32 Language
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Language");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Options Options
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Options>(_instance, "Options", NetOffice.PublisherApi.Options.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16), ProxyResult]
        public static object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string Path
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string PathSeparator
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "PathSeparator");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string ProductCode
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ProductCode");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool PrintPreview
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "PrintPreview");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "PrintPreview", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool ScreenUpdating
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ScreenUpdating");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ScreenUpdating", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Selection Selection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Selection>(_instance, "Selection", NetOffice.PublisherApi.Selection.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool SnapToGuides
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "SnapToGuides");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "SnapToGuides", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool SnapToObjects
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "SnapToObjects");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "SnapToObjects", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string TemplateFolderPath
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "TemplateFolderPath");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string Version
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.OfficeApi.OfficeDataSourceObject OfficeDataSourceObject
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeDataSourceObject>(_instance, "OfficeDataSourceObject", NetOffice.OfficeApi.OfficeDataSourceObject.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool WizardCatalogVisible
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "WizardCatalogVisible");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "WizardCatalogVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Documents Documents
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Documents>(_instance, "Documents", NetOffice.PublisherApi.Documents.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.WebOptions WebOptions
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebOptions>(_instance, "WebOptions", NetOffice.PublisherApi.WebOptions.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.InstalledPrinters InstalledPrinters
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.InstalledPrinters>(_instance, "InstalledPrinters", NetOffice.PublisherApi.InstalledPrinters.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(_instance, "MsoDebugOptions", NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool ValidateAddressVisible
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ValidateAddressVisible");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ValidateAddressVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool InsertBarcodeVisible
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "InsertBarcodeVisible");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "InsertBarcodeVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static string ShowFollowUpCustom
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ShowFollowUpCustom");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowFollowUpCustom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
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
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.OfficeApi.IAssistance Assistance
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(_instance, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.CaptionStyles CaptionStyles
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CaptionStyles>(_instance, "CaptionStyles", NetOffice.PublisherApi.CaptionStyles.LateBindingApiWrapperType);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="dir">string dir</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static void ChangeFileOpenDirectory(string dir)
        {
            Factory.ExecuteMethod(_instance, "ChangeFileOpenDirectory", dir);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="helpType">NetOffice.PublisherApi.Enums.PbHelpType helpType</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static void Help(NetOffice.PublisherApi.Enums.PbHelpType helpType)
        {
            Factory.ExecuteMethod(_instance, "Help", helpType);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="_object">object object</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static bool IsValidObject(object _object)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "IsValidObject", _object);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
        /// <param name="design">optional Int32 Design = -1</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document NewDocument(object wizard, object design)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "NewDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, wizard, design);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document NewDocument()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "NewDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document NewDocument(object wizard)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "NewDocument", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, wizard);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="readOnly">optional bool ReadOnly = false</param>
        /// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
        /// <param name="saveChanges">optional NetOffice.PublisherApi.Enums.PbSaveOptions SaveChanges = 1</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document Open(string filename, object readOnly, object addToRecentFiles, object saveChanges)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename, readOnly, addToRecentFiles, saveChanges);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document Open(string filename)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="readOnly">optional bool ReadOnly = false</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document Open(string filename, object readOnly)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename, readOnly);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="readOnly">optional bool ReadOnly = false</param>
        /// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static NetOffice.PublisherApi.Document Open(string filename, object readOnly, object addToRecentFiles)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(_instance, "Open", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, filename, readOnly, addToRecentFiles);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static void Quit()
        {
            Factory.ExecuteMethod(_instance, "Quit");
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static void LaunchWebService()
        {
            Factory.ExecuteMethod(_instance, "LaunchWebService");
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single CentimetersToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "CentimetersToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single EmusToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "EmusToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single InchesToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "InchesToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single LinesToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "LinesToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single MillimetersToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "MillimetersToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PicasToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PicasToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PixelsToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PixelsToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single TwipsToPoints(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "TwipsToPoints", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToCentimeters(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToCentimeters", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToEmus(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToEmus", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToInches(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToInches", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToLines(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToLines", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToMillimeters(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToMillimeters", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToPicas(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToPicas", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToPixels(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToPixels", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="value">Single value</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static Single PointsToTwips(Single value)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToTwips", value);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="wizard">optional NetOffice.PublisherApi.Enums.PbWizard Wizard = 0</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static void ShowWizardCatalog(object wizard)
        {
            Factory.ExecuteMethod(_instance, "ShowWizardCatalog", wizard);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public static void ShowWizardCatalog()
        {
            Factory.ExecuteMethod(_instance, "ShowWizardCatalog");
        }

        #endregion
    }
}
