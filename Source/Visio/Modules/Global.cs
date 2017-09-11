using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using System.Reflection;

namespace NetOffice.VisioApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Visio 11,12,14,15,16
    ///</summary>
    [SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(VisioApi.Application))]
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVDocument ActiveDocument
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(_instance, "ActiveDocument");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVPage ActivePage
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(_instance, "ActivePage");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVWindow ActiveWindow
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(_instance, "ActiveWindow");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVApplication Application
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(_instance, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVDocuments Documents
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocuments>(_instance, "Documents");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 ObjectType
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ObjectType");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 OnDataChangeDelay
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "OnDataChangeDelay");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "OnDataChangeDelay", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 ProcessID
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "ProcessID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 ScreenUpdating
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ScreenUpdating");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ScreenUpdating", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 Stat
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "Stat");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string Version
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Int16 WindowHandle
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "WindowHandle");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVWindows Windows
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindows>(_instance, "Windows");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 Language
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Language");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Int16 IsVisio16
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "IsVisio16");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Int16 IsVisio32
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "IsVisio32");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 WindowHandle32
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "WindowHandle32");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Int16 InstanceHandle
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "InstanceHandle");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 InstanceHandle32
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "InstanceHandle32");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVUIObject BuiltInMenus
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(_instance, "BuiltInMenus");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fIgnored">Int16 fIgnored</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.VisioApi.IVUIObject get_BuiltInToolbars(Int16 fIgnored)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(_instance, "BuiltInToolbars", NetOffice.VisioApi.IVUIObject.LateBindingApiWrapperType, fIgnored);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_BuiltInToolbars
        /// </summary>
        /// <param name="fIgnored">Int16 fIgnored</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_BuiltInToolbars")]
        public static NetOffice.VisioApi.IVUIObject BuiltInToolbars(Int16 fIgnored)
        {
            return get_BuiltInToolbars(fIgnored);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVUIObject CustomMenus
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(_instance, "CustomMenus");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string CustomMenusFile
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "CustomMenusFile");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "CustomMenusFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVUIObject CustomToolbars
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(_instance, "CustomToolbars");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string CustomToolbarsFile
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "CustomToolbarsFile");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "CustomToolbarsFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string AddonPaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "AddonPaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "AddonPaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string DrawingPaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "DrawingPaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DrawingPaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string FilterPaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "FilterPaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "FilterPaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string HelpPaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "HelpPaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "HelpPaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string StartupPaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "StartupPaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "StartupPaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string StencilPaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "StencilPaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "StencilPaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string TemplatePaths
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "TemplatePaths");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "TemplatePaths", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 PromptForSummary
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "PromptForSummary");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "PromptForSummary", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVAddons Addons
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVAddons>(_instance, "Addons");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string ProfileName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ProfileName");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="eventSeqNum">Int32 eventSeqNum</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string get_EventInfo(Int32 eventSeqNum)
        {
            return Factory.ExecuteStringPropertyGet(_instance, "EventInfo", eventSeqNum);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_EventInfo
        /// </summary>
        /// <param name="eventSeqNum">Int32 eventSeqNum</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_EventInfo")]
        public static string EventInfo(Int32 eventSeqNum)
        {
            return get_EventInfo(eventSeqNum);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVEventList EventList
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(_instance, "EventList");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 PersistsEvents
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "PersistsEvents");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 Active
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "Active");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 DeferRecalc
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "DeferRecalc");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DeferRecalc", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 AlertResponse
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "AlertResponse");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "AlertResponse", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 ShowProgress
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ShowProgress");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowProgress", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public static object Vbe
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "Vbe");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Int16 ShowMenus
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ShowMenus");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowMenus", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Int16 ToolbarStyle
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ToolbarStyle");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ToolbarStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 ShowStatusBar
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ShowStatusBar");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowStatusBar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 EventsEnabled
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "EventsEnabled");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "EventsEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string Path
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 TraceFlags
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "TraceFlags");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "TraceFlags", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 ShowToolbar
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "ShowToolbar");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowToolbar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool LiveDynamics
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "LiveDynamics");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "LiveDynamics", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool AutoLayout
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "AutoLayout");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "AutoLayout", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string CommandLine
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "CommandLine");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool IsUndoingOrRedoing
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "IsUndoingOrRedoing");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 CurrentScope
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "CurrentScope");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="nCmdID">Int32 nCmdID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static bool get_IsInScope(Int32 nCmdID)
        {
            return Factory.ExecuteBoolPropertyGet(_instance, "IsInScope", nCmdID);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_IsInScope
        /// </summary>
        /// <param name="nCmdID">Int32 nCmdID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_IsInScope")]
        public static bool IsInScope(Int32 nCmdID)
        {
            return get_IsInScope(nCmdID);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static object old_Addins
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "old_Addins");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static string ProductName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ProductName");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool UndoEnabled
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "UndoEnabled");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "UndoEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool ShowChanges
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "ShowChanges");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "ShowChanges", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 TypelibMajorVersion
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "TypelibMajorVersion");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 TypelibMinorVersion
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "TypelibMinorVersion");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int16 AutoRecoverInterval
        {
            get
            {
                return Factory.ExecuteInt16PropertyGet(_instance, "AutoRecoverInterval");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "AutoRecoverInterval", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool InhibitSelectChange
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "InhibitSelectChange");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "InhibitSelectChange", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static String[] AvailablePrinters
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = (object)Invoker.PropertyGet(_instance, "AvailablePrinters", paramsArray);
                return (String[])returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public static object CommandBars
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "CommandBars");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 Build
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "Build");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public static object COMAddIns
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "COMAddIns");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static object DefaultPageUnits
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(_instance, "DefaultPageUnits");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(_instance, "DefaultPageUnits", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static object DefaultTextUnits
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(_instance, "DefaultTextUnits");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(_instance, "DefaultTextUnits", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static object DefaultAngleUnits
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(_instance, "DefaultAngleUnits");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(_instance, "DefaultAngleUnits", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static object DefaultDurationUnits
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(_instance, "DefaultDurationUnits");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(_instance, "DefaultDurationUnits", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 FullBuild
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "FullBuild");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static bool VBAEnabled
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "VBAEnabled");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static NetOffice.VisioApi.Enums.VisZoomBehavior DefaultZoomBehavior
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisZoomBehavior>(_instance, "DefaultZoomBehavior");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "DefaultZoomBehavior", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        public static stdole.Font DialogFont
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(_instance, "DialogFont", paramsArray);
                return returnItem as stdole.Font;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 LanguageHelp
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(_instance, "LanguageHelp");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVWindow Window
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(_instance, "Window");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public static object ConnectorToolDataObject
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "ConnectorToolDataObject");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.VisioApi.IVApplicationSettings Settings
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplicationSettings>(_instance, "Settings");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public static object SaveAsWebObject
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "SaveAsWebObject");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static object MsoDebugOptions
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "MsoDebugOptions");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public static string MyShapesPath
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "MyShapesPath");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "MyShapesPath", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16), ProxyResult]
        public static object DefaultRectangleDataObject
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "DefaultRectangleDataObject");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public static bool DataFeaturesEnabled
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DataFeaturesEnabled");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16), ProxyResult]
        public static object LanguageSettings
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "LanguageSettings");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16), ProxyResult]
        public static object Assistance
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "Assistance");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static bool DeferRelationshipRecalc
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "DeferRelationshipRecalc");
            }
            set
            {
                Factory.ExecuteValuePropertySet(_instance, "DeferRelationshipRecalc", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static NetOffice.VisioApi.Enums.VisEdition CurrentEdition
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisEdition>(_instance, "CurrentEdition");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static Int64 InstanceHandle64
        {
            get
            {
                return Factory.ExecuteInt64PropertyGet(_instance, "InstanceHandle64");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void Quit()
        {
            Factory.ExecuteMethod(_instance, "Quit");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void Redo()
        {
            Factory.ExecuteMethod(_instance, "Redo");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void Undo()
        {
            Factory.ExecuteMethod(_instance, "Undo");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="menusObject">NetOffice.VisioApi.IVUIObject menusObject</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void SetCustomMenus(NetOffice.VisioApi.IVUIObject menusObject)
        {
            Factory.ExecuteMethod(_instance, "SetCustomMenus", menusObject);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void ClearCustomMenus()
        {
            Factory.ExecuteMethod(_instance, "ClearCustomMenus");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="toolbarsObject">NetOffice.VisioApi.IVUIObject toolbarsObject</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void SetCustomToolbars(NetOffice.VisioApi.IVUIObject toolbarsObject)
        {
            Factory.ExecuteMethod(_instance, "SetCustomToolbars", toolbarsObject);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void ClearCustomToolbars()
        {
            Factory.ExecuteMethod(_instance, "ClearCustomToolbars");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void SaveWorkspaceAs(string fileName)
        {
            Factory.ExecuteMethod(_instance, "SaveWorkspaceAs", fileName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="commandID">Int16 commandID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void DoCmd(Int16 commandID)
        {
            Factory.ExecuteMethod(_instance, "DoCmd", commandID);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="stringOrNumber">object stringOrNumber</param>
        /// <param name="unitsIn">object unitsIn</param>
        /// <param name="unitsOut">object unitsOut</param>
        /// <param name="format">string format</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string FormatResult(object stringOrNumber, object unitsIn, object unitsOut, string format)
        {
            return Factory.ExecuteStringMethodGet(_instance, "FormatResult", stringOrNumber, unitsIn, unitsOut, format);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="stringOrNumber">object stringOrNumber</param>
        /// <param name="unitsIn">object unitsIn</param>
        /// <param name="unitsOut">object unitsOut</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Double ConvertResult(object stringOrNumber, object unitsIn, object unitsOut)
        {
            return Factory.ExecuteDoubleMethodGet(_instance, "ConvertResult", stringOrNumber, unitsIn, unitsOut);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pathsString">string pathsString</param>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void EnumDirectories(string pathsString, out String[] nameArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            nameArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray(pathsString, (object)nameArray);
            Invoker.Method(_instance, "EnumDirectories", paramsArray, modifiers);
            nameArray = paramsArray[1] as String[];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void PurgeUndo()
        {
            Factory.ExecuteMethod(_instance, "PurgeUndo");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="contextString">string contextString</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 QueueMarkerEvent(string contextString)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "QueueMarkerEvent", contextString);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrUndoScopeName">string bstrUndoScopeName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static Int32 BeginUndoScope(string bstrUndoScopeName)
        {
            return Factory.ExecuteInt32MethodGet(_instance, "BeginUndoScope", bstrUndoScopeName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nScopeID">Int32 nScopeID</param>
        /// <param name="bCommit">bool bCommit</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void EndUndoScope(Int32 nScopeID, bool bCommit)
        {
            Factory.ExecuteMethod(_instance, "EndUndoScope", nScopeID, bCommit);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pUndoUnit">object pUndoUnit</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void AddUndoUnit(object pUndoUnit)
        {
            Factory.ExecuteMethod(_instance, "AddUndoUnit", pUndoUnit);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrScopeName">string bstrScopeName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void RenameCurrentScope(string bstrScopeName)
        {
            Factory.ExecuteMethod(_instance, "RenameCurrentScope", bstrScopeName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrHelpFileName">string bstrHelpFileName</param>
        /// <param name="command">Int32 command</param>
        /// <param name="data">Int32 data</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void InvokeHelp(string bstrHelpFileName, Int32 command, Int32 data)
        {
            Factory.ExecuteMethod(_instance, "InvokeHelp", bstrHelpFileName, command, data);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="uStateID">NetOffice.VisioApi.Enums.VisOnComponentEnterCodes uStateID</param>
        /// <param name="bEnter">bool bEnter</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static void OnComponentEnterState(NetOffice.VisioApi.Enums.VisOnComponentEnterCodes uStateID, bool bEnter)
        {
            Factory.ExecuteMethod(_instance, "OnComponentEnterState", uStateID, bEnter);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nWhichStatistic">Int32 nWhichStatistic</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static object GetUsageStatistic(Int32 nWhichStatistic)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "GetUsageStatistic", nWhichStatistic);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="stringOrNumber">object stringOrNumber</param>
        /// <param name="unitsIn">object unitsIn</param>
        /// <param name="unitsOut">object unitsOut</param>
        /// <param name="format">string format</param>
        /// <param name="langID">optional Int32 LangID = 0</param>
        /// <param name="calendarID">optional Int32 CalendarID = -1</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format, object langID, object calendarID)
        {
            return Factory.ExecuteStringMethodGet(_instance, "FormatResultEx", new object[] { stringOrNumber, unitsIn, unitsOut, format, langID, calendarID });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="stringOrNumber">object stringOrNumber</param>
        /// <param name="unitsIn">object unitsIn</param>
        /// <param name="unitsOut">object unitsOut</param>
        /// <param name="format">string format</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format)
        {
            return Factory.ExecuteStringMethodGet(_instance, "FormatResultEx", stringOrNumber, unitsIn, unitsOut, format);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="stringOrNumber">object stringOrNumber</param>
        /// <param name="unitsIn">object unitsIn</param>
        /// <param name="unitsOut">object unitsOut</param>
        /// <param name="format">string format</param>
        /// <param name="langID">optional Int32 LangID = 0</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public static string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format, object langID)
        {
            return Factory.ExecuteStringMethodGet(_instance, "FormatResultEx", new object[] { stringOrNumber, unitsIn, unitsOut, format, langID });
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="sourceAddOn">object sourceAddOn</param>
        /// <param name="targetDocument">NetOffice.VisioApi.IVDocument targetDocument</param>
        /// <param name="targetModes">NetOffice.VisioApi.Enums.VisRibbonXModes targetModes</param>
        /// <param name="friendlyName">string friendlyName</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static void RegisterRibbonX(object sourceAddOn, NetOffice.VisioApi.IVDocument targetDocument, NetOffice.VisioApi.Enums.VisRibbonXModes targetModes, string friendlyName)
        {
            Factory.ExecuteMethod(_instance, "RegisterRibbonX", sourceAddOn, targetDocument, targetModes, friendlyName);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="sourceAddOn">object sourceAddOn</param>
        /// <param name="targetDocument">NetOffice.VisioApi.IVDocument targetDocument</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static void UnregisterRibbonX(object sourceAddOn, NetOffice.VisioApi.IVDocument targetDocument)
        {
            Factory.ExecuteMethod(_instance, "UnregisterRibbonX", sourceAddOn, targetDocument);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="galleryName">string galleryName</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static bool GetPreviewEnabled(string galleryName)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "GetPreviewEnabled", galleryName);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="galleryName">string galleryName</param>
        /// <param name="onOrOff">bool onOrOff</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static void SetPreviewEnabled(string galleryName, bool onOrOff)
        {
            Factory.ExecuteMethod(_instance, "SetPreviewEnabled", galleryName, onOrOff);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="stencilType">NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType</param>
        /// <param name="measurementSystem">NetOffice.VisioApi.Enums.VisMeasurementSystem measurementSystem</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static string GetBuiltInStencilFile(NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType, NetOffice.VisioApi.Enums.VisMeasurementSystem measurementSystem)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetBuiltInStencilFile", stencilType, measurementSystem);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="stencilType">NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public static string GetCustomStencilFile(NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType)
        {
            return Factory.ExecuteStringMethodGet(_instance, "GetCustomStencilFile", stencilType);
        }

        #endregion
    }
}
