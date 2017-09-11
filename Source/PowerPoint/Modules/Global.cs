using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    ///</summary>
    [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(PowerPointApi.Application))]
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
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746387.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.Presentations Presentations
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Presentations>(_instance, "Presentations", NetOffice.PowerPointApi.Presentations.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746218.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.DocumentWindows Windows
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DocumentWindows>(_instance, "Windows", NetOffice.PowerPointApi.DocumentWindows.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.PowerPointApi.PPDialogs Dialogs
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PPDialogs>(_instance, "Dialogs", NetOffice.PowerPointApi.PPDialogs.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745295.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.DocumentWindow ActiveWindow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DocumentWindow>(_instance, "ActiveWindow", NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744912.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.Presentation ActivePresentation
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Presentation>(_instance, "ActivePresentation", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744816.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.SlideShowWindows SlideShowWindows
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowWindows>(_instance, "SlideShowWindows", NetOffice.PowerPointApi.SlideShowWindows.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744604.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.CommandBars CommandBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(_instance, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745905.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string Path
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746231.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746732.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string Caption
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Caption");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Caption", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Assistant Assistant
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(_instance, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.FileSearch FileSearch
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(_instance, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.IFind FileFind
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IFind>(_instance, "FileFind", NetOffice.OfficeApi.IFind.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746539.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string Build
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Build");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746225.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string Version
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744908.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string OperatingSystem
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "OperatingSystem");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744770.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string ActivePrinter
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ActivePrinter");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744619.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Creator
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744952.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.AddIns AddIns
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.AddIns>(_instance, "AddIns", NetOffice.PowerPointApi.AddIns.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744521.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(_instance, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745458.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static Single Left
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(_instance, "Left");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746410.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static Single Top
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(_instance, "Top");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746076.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static Single Width
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(_instance, "Width");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744702.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static Single Height
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(_instance, "Height");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744049.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.Enums.PpWindowState WindowState
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpWindowState>(_instance, "WindowState");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "WindowState", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745566.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoTriState Visible
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(_instance, "Visible");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 HWND
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "HWND");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745671.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoTriState Active
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(_instance, "Active");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.AnswerWizard AnswerWizard
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(_instance, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746702.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.COMAddIns COMAddIns
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(_instance, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744276.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static string ProductCode
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ProductCode");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.DefaultWebOptions DefaultWebOptions
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DefaultWebOptions>(_instance, "DefaultWebOptions", NetOffice.PowerPointApi.DefaultWebOptions.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745687.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.LanguageSettings LanguageSettings
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(_instance, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(_instance, "MsoDebugOptions", NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745929.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoTriState ShowWindowsInTaskbar
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(_instance, "ShowWindowsInTaskbar");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "ShowWindowsInTaskbar", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.PowerPointApi.Marker Marker
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Marker>(_instance, "Marker", NetOffice.PowerPointApi.Marker.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744258.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744016.aspx </remarks>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(_instance, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Alias for get_FileDialog
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744016.aspx </remarks>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16), Redirect("get_FileDialog")]
        public static NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
        {
            return get_FileDialog(type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746016.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoTriState DisplayGridLines
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(_instance, "DisplayGridLines");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "DisplayGridLines", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745661.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745695.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.NewFile NewPresentation
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(_instance, "NewPresentation", NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746503.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.Enums.PpAlertLevel DisplayAlerts
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpAlertLevel>(_instance, "DisplayAlerts");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "DisplayAlerts", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745925.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoTriState ShowStartupDialog
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(_instance, "ShowStartupDialog");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "ShowStartupDialog", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744774.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.AutoCorrect AutoCorrect
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.AutoCorrect>(_instance, "AutoCorrect", NetOffice.PowerPointApi.AutoCorrect.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744854.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.PowerPointApi.Options Options
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Options>(_instance, "Options", NetOffice.PowerPointApi.Options.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744758.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public static bool DisplayDocumentInformationPanel
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DisplayDocumentInformationPanel");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DisplayDocumentInformationPanel", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743833.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.IAssistance Assistance
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(_instance, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745260.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public static Int32 ActiveEncryptionSession
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "ActiveEncryptionSession");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744365.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.PowerPointApi.FileConverters FileConverters
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.FileConverters>(_instance, "FileConverters", NetOffice.PowerPointApi.FileConverters.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743963.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtLayouts>(_instance, "SmartArtLayouts", NetOffice.OfficeApi.SmartArtLayouts.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745345.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtQuickStyles>(_instance, "SmartArtQuickStyles", NetOffice.OfficeApi.SmartArtQuickStyles.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745159.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.OfficeApi.SmartArtColors SmartArtColors
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtColors>(_instance, "SmartArtColors", NetOffice.OfficeApi.SmartArtColors.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744225.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.PowerPointApi.ProtectedViewWindows ProtectedViewWindows
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ProtectedViewWindows>(_instance, "ProtectedViewWindows", NetOffice.PowerPointApi.ProtectedViewWindows.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746155.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.PowerPointApi.ProtectedViewWindow ActiveProtectedViewWindow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ProtectedViewWindow>(_instance, "ActiveProtectedViewWindow", NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746169.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static bool IsSandboxed
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "IsSandboxed");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.PowerPointApi.ResampleMediaTasks ResampleMediaTasks
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ResampleMediaTasks>(_instance, "ResampleMediaTasks", NetOffice.PowerPointApi.ResampleMediaTasks.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745623.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileValidationMode>(_instance, "FileValidation");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "FileValidation", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229713.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public static bool ChartDataPointTrack
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ChartDataPointTrack");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ChartDataPointTrack", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228516.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoTriState DisplayGuides
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(_instance, "DisplayGuides");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "DisplayGuides", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745709.aspx </remarks>
        /// <param name="helpFile">optional string HelpFile = vbappt9.chm</param>
        /// <param name="contextID">optional Int32 ContextID = 0</param>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void Help(object helpFile, object contextID)
        {
            Factory.ExecuteMethod(_instance, "Help", helpFile, contextID);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745709.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void Help()
        {
            Factory.ExecuteMethod(_instance, "Help");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745709.aspx </remarks>
        /// <param name="helpFile">optional string HelpFile = vbappt9.chm</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void Help(object helpFile)
        {
            Factory.ExecuteMethod(_instance, "Help", helpFile);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746388.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void Quit()
        {
            Factory.ExecuteMethod(_instance, "Quit");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744221.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="safeArrayOfParams">optional object[] safeArrayOfParams</param>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object[] safeArrayOfParams)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macroName, safeArrayOfParams);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744221.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macroName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9
        /// </summary>
        /// <param name="type">NetOffice.PowerPointApi.Enums.PpFileDialogType type</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 9)]
        public static NetOffice.PowerPointApi.FileDialog FileDialog(NetOffice.PowerPointApi.Enums.PpFileDialogType type)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.FileDialog>(_instance, "FileDialog", NetOffice.PowerPointApi.FileDialog.LateBindingApiWrapperType, type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pWindow">NetOffice.PowerPointApi.DocumentWindow pWindow</param>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void LaunchSpelling(NetOffice.PowerPointApi.DocumentWindow pWindow)
        {
            Factory.ExecuteMethod(_instance, "LaunchSpelling", pWindow);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745072.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void Activate()
        {
            Factory.ExecuteMethod(_instance, "Activate");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="option">Int32 option</param>
        /// <param name="persist">optional bool Persist = false</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static bool GetOptionFlag(Int32 option, object persist)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "GetOptionFlag", option, persist);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="option">Int32 option</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static bool GetOptionFlag(Int32 option)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "GetOptionFlag", option);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="option">Int32 option</param>
        /// <param name="state">bool state</param>
        /// <param name="persist">optional bool Persist = false</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void SetOptionFlag(Int32 option, bool state, object persist)
        {
            Factory.ExecuteMethod(_instance, "SetOptionFlag", option, state, persist);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="option">Int32 option</param>
        /// <param name="state">bool state</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public static void SetOptionFlag(Int32 option, bool state)
        {
            Factory.ExecuteMethod(_instance, "SetOptionFlag", option, state);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.PowerPointApi.Enums.PpFileDialogType type</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static object PPFileDialog(NetOffice.PowerPointApi.Enums.PpFileDialogType type)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "PPFileDialog", type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="marker">Int32 marker</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public static void SetPerfMarker(Int32 marker)
        {
            Factory.ExecuteMethod(_instance, "SetPerfMarker", marker);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 12, 14, 15, 16
        /// </summary>
        /// <param name="slideLibraryUrl">string slideLibraryUrl</param>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public static void LaunchPublishSlidesDialog(string slideLibraryUrl)
        {
            Factory.ExecuteMethod(_instance, "LaunchPublishSlidesDialog", slideLibraryUrl);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 12, 14, 15, 16
        /// </summary>
        /// <param name="slideUrls">object slideUrls</param>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public static void LaunchSendToPPTDialog(object slideUrls)
        {
            Factory.ExecuteMethod(_instance, "LaunchSendToPPTDialog", slideUrls);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745395.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public static void StartNewUndoEntry()
        {
            Factory.ExecuteMethod(_instance, "StartNewUndoEntry");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229517.aspx </remarks>
        /// <param name="themeFileName">string themeFileName</param>
        [SupportByVersion("PowerPoint", 15, 16)]
        public static NetOffice.PowerPointApi.Theme OpenThemeFile(string themeFileName)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Theme>(_instance, "OpenThemeFile", NetOffice.PowerPointApi.Theme.LateBindingApiWrapperType, themeFileName);
        }

        #endregion
    }
}
