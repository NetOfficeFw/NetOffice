using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    ///</summary>
    [SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(WordApi.Application))]
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823254.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(_instance, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197825.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Creator
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191758.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public static object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845178.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821628.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Documents Documents
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Documents>(_instance, "Documents", NetOffice.WordApi.Documents.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822351.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Windows Windows
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Windows>(_instance, "Windows", NetOffice.WordApi.Windows.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837737.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document ActiveDocument
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(_instance, "ActiveDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845301.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Window ActiveWindow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(_instance, "ActiveWindow", NetOffice.WordApi.Window.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838682.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Selection Selection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Selection>(_instance, "Selection", NetOffice.WordApi.Selection.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822917.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public static object WordBasic
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "WordBasic");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195679.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.RecentFiles RecentFiles
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RecentFiles>(_instance, "RecentFiles", NetOffice.WordApi.RecentFiles.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845589.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Template NormalTemplate
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Template>(_instance, "NormalTemplate", NetOffice.WordApi.Template.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822391.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.System System
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.System>(_instance, "System", NetOffice.WordApi.System.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845308.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.AutoCorrect AutoCorrect
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrect>(_instance, "AutoCorrect", NetOffice.WordApi.AutoCorrect.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197817.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.FontNames FontNames
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(_instance, "FontNames", NetOffice.WordApi.FontNames.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196340.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.FontNames LandscapeFontNames
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(_instance, "LandscapeFontNames", NetOffice.WordApi.FontNames.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192201.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.FontNames PortraitFontNames
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(_instance, "PortraitFontNames", NetOffice.WordApi.FontNames.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840701.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Languages Languages
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Languages>(_instance, "Languages", NetOffice.WordApi.Languages.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Assistant Assistant
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(_instance, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821300.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Browser Browser
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Browser>(_instance, "Browser", NetOffice.WordApi.Browser.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823259.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.FileConverters FileConverters
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FileConverters>(_instance, "FileConverters", NetOffice.WordApi.FileConverters.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821659.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.MailingLabel MailingLabel
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailingLabel>(_instance, "MailingLabel", NetOffice.WordApi.MailingLabel.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191745.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Dialogs Dialogs
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Dialogs>(_instance, "Dialogs", NetOffice.WordApi.Dialogs.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838479.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.CaptionLabels CaptionLabels
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CaptionLabels>(_instance, "CaptionLabels", NetOffice.WordApi.CaptionLabels.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198063.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.AutoCaptions AutoCaptions
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCaptions>(_instance, "AutoCaptions", NetOffice.WordApi.AutoCaptions.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822986.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.AddIns AddIns
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AddIns>(_instance, "AddIns", NetOffice.WordApi.AddIns.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839544.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821519.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string Version
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197438.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198164.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839740.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Tasks Tasks
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tasks>(_instance, "Tasks", NetOffice.WordApi.Tasks.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool DisplayStatusBar
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DisplayStatusBar");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DisplayStatusBar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836086.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool SpecialMode
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "SpecialMode");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839688.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 UsableWidth
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "UsableWidth");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834606.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 UsableHeight
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "UsableHeight");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192165.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool MathCoprocessorAvailable
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "MathCoprocessorAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192426.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool MouseAvailable
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "MouseAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
        /// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static object get_International(NetOffice.WordApi.Enums.WdInternationalIndex index)
        {
            return Factory.ExecuteVariantPropertyGet(_instance, "International", index);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_International
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
        /// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_International")]
        public static object International(NetOffice.WordApi.Enums.WdInternationalIndex index)
        {
            return get_International(index);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839495.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string Build
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Build");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820850.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CapsLock
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "CapsLock");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845392.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool NumLock
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "NumLock");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834599.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string UserName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "UserName");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "UserName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844813.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string UserInitials
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "UserInitials");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "UserInitials", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193411.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string UserAddress
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "UserAddress");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "UserAddress", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835128.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public static object MacroContainer
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "MacroContainer");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838964.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool DisplayRecentFiles
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DisplayRecentFiles");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DisplayRecentFiles", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845623.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.CommandBars CommandBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(_instance, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word, object languageID)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(_instance, "SynonymInfo", NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType, word, languageID);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_SynonymInfo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_SynonymInfo")]
        public static NetOffice.WordApi.SynonymInfo SynonymInfo(string word, object languageID)
        {
            return get_SynonymInfo(word, languageID);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(_instance, "SynonymInfo", NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType, word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_SynonymInfo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_SynonymInfo")]
        public static NetOffice.WordApi.SynonymInfo SynonymInfo(string word)
        {
            return get_SynonymInfo(word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197234.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(_instance, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839412.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string DefaultSaveFormat
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "DefaultSaveFormat");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DefaultSaveFormat", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821102.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.ListGalleries ListGalleries
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListGalleries>(_instance, "ListGalleries", NetOffice.WordApi.ListGalleries.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821995.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string ActivePrinter
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ActivePrinter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ActivePrinter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821925.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Templates Templates
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Templates>(_instance, "Templates", NetOffice.WordApi.Templates.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822548.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public static object CustomizationContext
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "CustomizationContext");
            }
            set
            {
                Factory.ExecuteReferencePropertySet(_instance, "CustomizationContext", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197596.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.KeyBindings KeyBindings
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBindings>(_instance, "KeyBindings", NetOffice.WordApi.KeyBindings.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        /// <param name="commandParameter">optional object commandParameter</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeysBoundTo>(_instance, "KeysBoundTo", NetOffice.WordApi.KeysBoundTo.LateBindingApiWrapperType, keyCategory, command, commandParameter);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_KeysBoundTo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        /// <param name="commandParameter">optional object commandParameter</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_KeysBoundTo")]
        public static NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
        {
            return get_KeysBoundTo(keyCategory, command, commandParameter);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeysBoundTo>(_instance, "KeysBoundTo", NetOffice.WordApi.KeysBoundTo.LateBindingApiWrapperType, keyCategory, command);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_KeysBoundTo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_KeysBoundTo")]
        public static NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
        {
            return get_KeysBoundTo(keyCategory, command);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode, object keyCode2)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBinding>(_instance, "FindKey", NetOffice.WordApi.KeyBinding.LateBindingApiWrapperType, keyCode, keyCode2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_FindKey
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_FindKey")]
        public static NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode, object keyCode2)
        {
            return get_FindKey(keyCode, keyCode2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBinding>(_instance, "FindKey", NetOffice.WordApi.KeyBinding.LateBindingApiWrapperType, keyCode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_FindKey
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_FindKey")]
        public static NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode)
        {
            return get_FindKey(keyCode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196028.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192216.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string Path
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192367.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool DisplayScrollBars
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DisplayScrollBars");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DisplayScrollBars", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string StartupPath
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "StartupPath");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "StartupPath", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835146.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 BackgroundSavingStatus
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "BackgroundSavingStatus");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820962.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 BackgroundPrintingStatus
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "BackgroundPrintingStatus");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839318.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Left
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Left");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837463.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Top
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Top");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836284.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Width
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Width");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845159.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Height
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Height");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836388.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Enums.WdWindowState WindowState
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdWindowState>(_instance, "WindowState");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "WindowState", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192152.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool DisplayAutoCompleteTips
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DisplayAutoCompleteTips");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DisplayAutoCompleteTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822542.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Options Options
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Options>(_instance, "Options", NetOffice.WordApi.Options.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192373.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Enums.WdAlertLevel DisplayAlerts
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdAlertLevel>(_instance, "DisplayAlerts");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "DisplayAlerts", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191957.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Dictionaries CustomDictionaries
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Dictionaries>(_instance, "CustomDictionaries", NetOffice.WordApi.Dictionaries.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192616.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string PathSeparator
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "PathSeparator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845291.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string StatusBar
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "StatusBar");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "StatusBar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192800.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool MAPIAvailable
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "MAPIAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845182.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool DisplayScreenTips
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DisplayScreenTips");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DisplayScreenTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839294.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Enums.WdEnableCancelKey EnableCancelKey
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEnableCancelKey>(_instance, "EnableCancelKey");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "EnableCancelKey", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197424.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool UserControl
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "UserControl");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.FileSearch FileSearch
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(_instance, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838972.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Enums.WdMailSystem MailSystem
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailSystem>(_instance, "MailSystem");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string DefaultTableSeparator
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "DefaultTableSeparator");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DefaultTableSeparator", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839922.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool ShowVisualBasicEditor
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ShowVisualBasicEditor");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowVisualBasicEditor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839549.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string BrowseExtraFileTypes
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "BrowseExtraFileTypes");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "BrowseExtraFileTypes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
        /// <param name="_object">object object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static bool get_IsObjectValid(object _object)
        {
            return Factory.ExecuteBoolPropertyGet(_instance, "IsObjectValid", _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_IsObjectValid
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
        /// <param name="_object">object object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_IsObjectValid")]
        public static bool IsObjectValid(object _object)
        {
            return get_IsObjectValid(_object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194713.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.HangulHanjaConversionDictionaries HangulHanjaDictionaries
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HangulHanjaConversionDictionaries>(_instance, "HangulHanjaDictionaries", NetOffice.WordApi.HangulHanjaConversionDictionaries.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821986.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.MailMessage MailMessage
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMessage>(_instance, "MailMessage", NetOffice.WordApi.MailMessage.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840871.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool FocusInMailHeader
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "FocusInMailHeader");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192588.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.EmailOptions EmailOptions
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.EmailOptions>(_instance, "EmailOptions", NetOffice.WordApi.EmailOptions.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836711.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Enums.MsoLanguageID Language
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(_instance, "Language");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192831.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.COMAddIns COMAddIns
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(_instance, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192428.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckLanguage
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "CheckLanguage");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "CheckLanguage", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197161.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.LanguageSettings LanguageSettings
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(_instance, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static bool Dummy1
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "Dummy1");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.AnswerWizard AnswerWizard
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(_instance, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195192.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192776.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
        /// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(_instance, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, fileDialogType);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Alias for get_FileDialog
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
        /// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16), Redirect("get_FileDialog")]
        public static NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
        {
            return get_FileDialog(fileDialogType);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193382.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static string EmailTemplate
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "EmailTemplate");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "EmailTemplate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static bool ShowWindowsInTaskbar
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ShowWindowsInTaskbar");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowWindowsInTaskbar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193065.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.NewFile NewDocument
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(_instance, "NewDocument", NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840052.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static bool ShowStartupDialog
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ShowStartupDialog");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowStartupDialog", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192177.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.AutoCorrect AutoCorrectEmail
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrect>(_instance, "AutoCorrectEmail", NetOffice.WordApi.AutoCorrect.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845341.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.TaskPanes TaskPanes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TaskPanes>(_instance, "TaskPanes", NetOffice.WordApi.TaskPanes.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835491.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static bool DefaultLegalBlackline
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DefaultLegalBlackline");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DefaultLegalBlackline", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SmartTagRecognizers SmartTagRecognizers
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTagRecognizers>(_instance, "SmartTagRecognizers", NetOffice.WordApi.SmartTagRecognizers.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SmartTagTypes SmartTagTypes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTagTypes>(_instance, "SmartTagTypes", NetOffice.WordApi.SmartTagTypes.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839771.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.XMLNamespaces XMLNamespaces
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNamespaces>(_instance, "XMLNamespaces", NetOffice.WordApi.XMLNamespaces.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196679.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public static bool ArbitraryXMLSupportAvailable
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ArbitraryXMLSupportAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string BuildFull
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "BuildFull");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string BuildFeatureCrew
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "BuildFeatureCrew");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192405.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Bibliography Bibliography
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bibliography>(_instance, "Bibliography", NetOffice.WordApi.Bibliography.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191727.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static bool ShowStylePreviews
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ShowStylePreviews");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowStylePreviews", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845435.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static bool RestrictLinkedStyles
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "RestrictLinkedStyles");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "RestrictLinkedStyles", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837322.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.OMathAutoCorrect OMathAutoCorrect
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathAutoCorrect>(_instance, "OMathAutoCorrect", NetOffice.WordApi.OMathAutoCorrect.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836074.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
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
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197133.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.IAssistance Assistance
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(_instance, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192620.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static bool OpenAttachmentsInFullScreen
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "OpenAttachmentsInFullScreen");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "OpenAttachmentsInFullScreen", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836063.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static Int32 ActiveEncryptionSession
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "ActiveEncryptionSession");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194203.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static bool DontResetInsertionPointProperties
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DontResetInsertionPointProperties");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DontResetInsertionPointProperties", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839192.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtLayouts>(_instance, "SmartArtLayouts", NetOffice.OfficeApi.SmartArtLayouts.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194982.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtQuickStyles>(_instance, "SmartArtQuickStyles", NetOffice.OfficeApi.SmartArtQuickStyles.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839505.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.OfficeApi.SmartArtColors SmartArtColors
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtColors>(_instance, "SmartArtColors", NetOffice.OfficeApi.SmartArtColors.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838675.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.WordApi.UndoRecord UndoRecord
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.UndoRecord>(_instance, "UndoRecord", NetOffice.WordApi.UndoRecord.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191978.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.OfficeApi.PickerDialog PickerDialog
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerDialog>(_instance, "PickerDialog", NetOffice.OfficeApi.PickerDialog.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839925.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.WordApi.ProtectedViewWindows ProtectedViewWindows
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProtectedViewWindows>(_instance, "ProtectedViewWindows", NetOffice.WordApi.ProtectedViewWindows.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192773.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static NetOffice.WordApi.ProtectedViewWindow ActiveProtectedViewWindow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProtectedViewWindow>(_instance, "ActiveProtectedViewWindow", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845787.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public static bool IsSandboxed
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "IsSandboxed");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193078.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232091.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
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
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232207.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public static bool ShowAnimation
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ShowAnimation");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowAnimation", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="originalFormat">optional object originalFormat</param>
        /// <param name="routeDocument">optional object routeDocument</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Quit(object saveChanges, object originalFormat, object routeDocument)
        {
            Factory.ExecuteMethod(_instance, "Quit", saveChanges, originalFormat, routeDocument);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Quit()
        {
            Factory.ExecuteMethod(_instance, "Quit");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Quit(object saveChanges)
        {
            Factory.ExecuteMethod(_instance, "Quit", saveChanges);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="originalFormat">optional object originalFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Quit(object saveChanges, object originalFormat)
        {
            Factory.ExecuteMethod(_instance, "Quit", saveChanges, originalFormat);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193095.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ScreenRefresh()
        {
            Factory.ExecuteMethod(_instance, "ScreenRefresh");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld()
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", background);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", background, append);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", background, append, range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", background, append, range, outputFileName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
        {
            Factory.ExecuteMethod(_instance, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839803.aspx </remarks>
        /// <param name="name">string name</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void LookupNameProperties(string name)
        {
            Factory.ExecuteMethod(_instance, "LookupNameProperties", name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192415.aspx </remarks>
        /// <param name="unavailableFont">string unavailableFont</param>
        /// <param name="substituteFont">string substituteFont</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void SubstituteFont(string unavailableFont, string substituteFont)
        {
            Factory.ExecuteMethod(_instance, "SubstituteFont", unavailableFont, substituteFont);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
        /// <param name="times">optional object times</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool Repeat(object times)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "Repeat", times);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool Repeat()
        {
            return Factory.ExecuteBoolMethodGet(_instance, "Repeat");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845561.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDEExecute(Int32 channel, string command)
        {
            Factory.ExecuteMethod(_instance, "DDEExecute", channel, command);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837295.aspx </remarks>
        /// <param name="app">string app</param>
        /// <param name="topic">string topic</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 DDEInitiate(string app, string topic)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "DDEInitiate", app, topic);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837201.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="item">string item</param>
        /// <param name="data">string data</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDEPoke(Int32 channel, string item, string data)
        {
            Factory.ExecuteMethod(_instance, "DDEPoke", channel, item, data);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837546.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="item">string item</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string DDERequest(Int32 channel, string item)
        {
            return Factory.ExecuteStringMethodGet(_instance, "DDERequest", channel, item);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837904.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDETerminate(Int32 channel)
        {
            Factory.ExecuteMethod(_instance, "DDETerminate", channel);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192053.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void DDETerminateAll()
        {
            Factory.ExecuteMethod(_instance, "DDETerminateAll");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3, object arg4)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "BuildKeyCode", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "BuildKeyCode", arg1);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "BuildKeyCode", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "BuildKeyCode", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string KeyString(Int32 keyCode, object keyCode2)
        {
            return Factory.ExecuteStringMethodGet(_instance, "KeyString", keyCode, keyCode2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string KeyString(Int32 keyCode)
        {
            return Factory.ExecuteStringMethodGet(_instance, "KeyString", keyCode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835492.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="destination">string destination</param>
        /// <param name="name">string name</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void OrganizerCopy(string source, string destination, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
        {
            Factory.ExecuteMethod(_instance, "OrganizerCopy", source, destination, name, _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194744.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="name">string name</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void OrganizerDelete(string source, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
        {
            Factory.ExecuteMethod(_instance, "OrganizerDelete", source, name, _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836140.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void OrganizerRename(string source, string name, string newName, NetOffice.WordApi.Enums.WdOrganizerObject _object)
        {
            Factory.ExecuteMethod(_instance, "OrganizerRename", source, name, newName, _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823266.aspx </remarks>
        /// <param name="tagID">String[] tagID</param>
        /// <param name="value">String[] value</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void AddAddress(String[] tagID, String[] value)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)tagID, (object)value);
            Invoker.Method(_instance, "AddAddress", paramsArray); ;
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        /// <param name="checkNamesDialog">optional object checkNamesDialog</param>
        /// <param name="recentAddressesChoice">optional object recentAddressesChoice</param>
        /// <param name="updateRecentAddresses">optional object updateRecentAddresses</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice, object updateRecentAddresses)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice, updateRecentAddresses });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress()
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", name, addressProperties);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties, object useAutoText)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", name, addressProperties, useAutoText);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", name, addressProperties, useAutoText, displaySelectDialog);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        /// <param name="checkNamesDialog">optional object checkNamesDialog</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        /// <param name="checkNamesDialog">optional object checkNamesDialog</param>
        /// <param name="recentAddressesChoice">optional object recentAddressesChoice</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194798.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckGrammar(string _string)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckGrammar", _string);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        /// <param name="customDictionary10">optional object customDictionary10</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", word, customDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", word, customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", word, customDictionary, ignoreUppercase, mainDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822681.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ResetIgnoreAll()
        {
            Factory.ExecuteMethod(_instance, "ResetIgnoreAll");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        /// <param name="customDictionary10">optional object customDictionary10</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word, customDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word, customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word, customDictionary, ignoreUppercase, mainDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(_instance, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838545.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void GoBack()
        {
            Factory.ExecuteMethod(_instance, "GoBack");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841057.aspx </remarks>
        /// <param name="helpType">object helpType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Help(object helpType)
        {
            Factory.ExecuteMethod(_instance, "Help", helpType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194337.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void AutomaticChange()
        {
            Factory.ExecuteMethod(_instance, "AutomaticChange");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839095.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ShowMe()
        {
            Factory.ExecuteMethod(_instance, "ShowMe");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821932.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void HelpTool()
        {
            Factory.ExecuteMethod(_instance, "HelpTool");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845336.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.Window NewWindow()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Window>(_instance, "NewWindow", NetOffice.WordApi.Window.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194509.aspx </remarks>
        /// <param name="listAllCommands">bool listAllCommands</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ListCommands(bool listAllCommands)
        {
            Factory.ExecuteMethod(_instance, "ListCommands", listAllCommands);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834517.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ShowClipboard()
        {
            Factory.ExecuteMethod(_instance, "ShowClipboard");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
        /// <param name="when">object when</param>
        /// <param name="name">string name</param>
        /// <param name="tolerance">optional object tolerance</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void OnTime(object when, string name, object tolerance)
        {
            Factory.ExecuteMethod(_instance, "OnTime", when, name, tolerance);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
        /// <param name="when">object when</param>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void OnTime(object when, string name)
        {
            Factory.ExecuteMethod(_instance, "OnTime", when, name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837154.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void NextLetter()
        {
            Factory.ExecuteMethod(_instance, "NextLetter");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        /// <param name="userPassword">optional object userPassword</param>
        /// <param name="volumePassword">optional object volumePassword</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int16 MountVolume(string zone, string server, string volume, object user, object userPassword, object volumePassword)
        {
            return Factory.ExecuteInt16MethodGet(_instance, "MountVolume", new object[] { zone, server, volume, user, userPassword, volumePassword });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int16 MountVolume(string zone, string server, string volume)
        {
            return Factory.ExecuteInt16MethodGet(_instance, "MountVolume", zone, server, volume);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int16 MountVolume(string zone, string server, string volume, object user)
        {
            return Factory.ExecuteInt16MethodGet(_instance, "MountVolume", zone, server, volume, user);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        /// <param name="userPassword">optional object userPassword</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int16 MountVolume(string zone, string server, string volume, object user, object userPassword)
        {
            return Factory.ExecuteInt16MethodGet(_instance, "MountVolume", new object[] { zone, server, volume, user, userPassword });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844818.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string CleanString(string _string)
        {
            return Factory.ExecuteStringMethodGet(_instance, "CleanString", _string);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void SendFax()
        {
            Factory.ExecuteMethod(_instance, "SendFax");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835219.aspx </remarks>
        /// <param name="path">string path</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ChangeFileOpenDirectory(string path)
        {
            Factory.ExecuteMethod(_instance, "ChangeFileOpenDirectory", path);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macroName">string macroName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void RunOld(string macroName)
        {
            Factory.ExecuteMethod(_instance, "RunOld", macroName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196922.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void GoForward()
        {
            Factory.ExecuteMethod(_instance, "GoForward");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844914.aspx </remarks>
        /// <param name="left">Int32 left</param>
        /// <param name="top">Int32 top</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Move(Int32 left, Int32 top)
        {
            Factory.ExecuteMethod(_instance, "Move", left, top);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197452.aspx </remarks>
        /// <param name="width">Int32 width</param>
        /// <param name="height">Int32 height</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Resize(Int32 width, Int32 height)
        {
            Factory.ExecuteMethod(_instance, "Resize", width, height);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197549.aspx </remarks>
        /// <param name="inches">Single inches</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single InchesToPoints(Single inches)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "InchesToPoints", inches);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838159.aspx </remarks>
        /// <param name="centimeters">Single centimeters</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single CentimetersToPoints(Single centimeters)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "CentimetersToPoints", centimeters);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845767.aspx </remarks>
        /// <param name="millimeters">Single millimeters</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single MillimetersToPoints(Single millimeters)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "MillimetersToPoints", millimeters);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840225.aspx </remarks>
        /// <param name="picas">Single picas</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PicasToPoints(Single picas)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PicasToPoints", picas);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840343.aspx </remarks>
        /// <param name="lines">Single lines</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single LinesToPoints(Single lines)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "LinesToPoints", lines);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838268.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToInches(Single points)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToInches", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195052.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToCentimeters(Single points)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToCentimeters", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836929.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToMillimeters(Single points)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToMillimeters", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193434.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToPicas(Single points)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToPicas", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822110.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToLines(Single points)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToLines", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821351.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void Activate()
        {
            Factory.ExecuteMethod(_instance, "Activate");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
        /// <param name="points">Single points</param>
        /// <param name="fVertical">optional object fVertical</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToPixels(Single points, object fVertical)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToPixels", points, fVertical);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
        /// <param name="points">Single points</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PointsToPixels(Single points)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PointsToPixels", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
        /// <param name="pixels">Single pixels</param>
        /// <param name="fVertical">optional object fVertical</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PixelsToPoints(Single pixels, object fVertical)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PixelsToPoints", pixels, fVertical);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
        /// <param name="pixels">Single pixels</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Single PixelsToPoints(Single pixels)
        {
            return Factory.ExecuteSingleMethodGet(_instance, "PixelsToPoints", pixels);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845662.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void KeyboardLatin()
        {
            Factory.ExecuteMethod(_instance, "KeyboardLatin");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196621.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void KeyboardBidi()
        {
            Factory.ExecuteMethod(_instance, "KeyboardBidi");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835971.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void ToggleKeyboard()
        {
            Factory.ExecuteMethod(_instance, "ToggleKeyboard");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
        /// <param name="langId">optional Int32 LangId = 0</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Keyboard(object langId)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "Keyboard", langId);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static Int32 Keyboard()
        {
            return Factory.ExecuteInt32MethodGet(_instance, "Keyboard");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193728.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string ProductCode()
        {
            return Factory.ExecuteStringMethodGet(_instance, "ProductCode");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840160.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.WordApi.DefaultWebOptions DefaultWebOptions()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DefaultWebOptions>(_instance, "DefaultWebOptions", NetOffice.WordApi.DefaultWebOptions.LateBindingApiWrapperType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="range">object range</param>
        /// <param name="cid">object cid</param>
        /// <param name="piCSE">object piCSE</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void DiscussionSupport(object range, object cid, object piCSE)
        {
            Factory.ExecuteMethod(_instance, "DiscussionSupport", range, cid, piCSE);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821531.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void SetDefaultTheme(string name, NetOffice.WordApi.Enums.WdDocumentMedium documentType)
        {
            Factory.ExecuteMethod(_instance, "SetDefaultTheme", name, documentType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834585.aspx </remarks>
        /// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static string GetDefaultTheme(NetOffice.WordApi.Enums.WdDocumentMedium documentType)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetDefaultTheme", documentType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        /// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut()
        {
            Factory.ExecuteMethod(_instance, "PrintOut");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", background);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", background, append);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", background, append, range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", background, append, range, outputFileName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
        {
            Factory.ExecuteMethod(_instance, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        /// <param name="varg28">optional object varg28</param>
        /// <param name="varg29">optional object varg29</param>
        /// <param name="varg30">optional object varg30</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29, object varg30)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macroName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macroName, varg1);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macroName, varg1, varg2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macroName, varg1, varg2, varg3);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        /// <param name="varg28">optional object varg28</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        /// <param name="varg28">optional object varg28</param>
        /// <param name="varg29">optional object varg29</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public static object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        /// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000()
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", background);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", background, append);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", background, append, range);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", background, append, range, outputFileName);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
        {
            Factory.ExecuteMethod(_instance, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public static bool Dummy2()
        {
            return Factory.ExecuteBoolMethodGet(_instance, "Dummy2");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838158.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public static void PutFocusInMailHeader()
        {
            Factory.ExecuteMethod(_instance, "PutFocusInMailHeader");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840673.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static void LoadMasterList(string fileName)
        {
            Factory.ExecuteMethod(_instance, "LoadMasterList", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        /// <param name="ignoreAllComparisonWarnings">optional bool IgnoreAllComparisonWarnings = false</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor, object ignoreAllComparisonWarnings)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor, ignoreAllComparisonWarnings });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination, granularity);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="originalAuthor">optional string OriginalAuthor = </param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        /// <param name="formatFrom">optional NetOffice.WordApi.Enums.WdMergeFormatFrom FormatFrom = 2</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor, object formatFrom)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor, formatFrom });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination, granularity);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="originalAuthor">optional string OriginalAuthor = </param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="originalAuthor">optional string OriginalAuthor = </param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public static NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(_instance, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="localDocument">NetOffice.WordApi.Document localDocument</param>
        /// <param name="serverDocument">NetOffice.WordApi.Document serverDocument</param>
        /// <param name="baseDocument">NetOffice.WordApi.Document baseDocument</param>
        /// <param name="favorSource">bool favorSource</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public static void ThreeWayMerge(NetOffice.WordApi.Document localDocument, NetOffice.WordApi.Document serverDocument, NetOffice.WordApi.Document baseDocument, bool favorSource)
        {
            Factory.ExecuteMethod(_instance, "ThreeWayMerge", localDocument, serverDocument, baseDocument, favorSource);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public static void Dummy4()
        {
            Factory.ExecuteMethod(_instance, "Dummy4");
        }

        #endregion
    }
}
