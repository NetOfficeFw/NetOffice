using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOAPPLICATION 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOAPPLICATION : COMObject
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
                    _type = typeof(LPVISIOAPPLICATION);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOAPPLICATION(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOAPPLICATION(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOAPPLICATION(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOAPPLICATION(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOAPPLICATION(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOAPPLICATION(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOAPPLICATION() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOAPPLICATION(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument ActiveDocument
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "ActiveDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPage ActivePage
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ActivePage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow ActiveWindow
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ActiveWindow");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocuments Documents
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocuments>(this, "Documents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 OnDataChangeDelay
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OnDataChangeDelay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDataChangeDelay", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ProcessID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ProcessID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ScreenUpdating
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ScreenUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 WindowHandle
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowHandle");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindows Windows
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindows>(this, "Windows");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Language
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Language");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 IsVisio16
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IsVisio16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 IsVisio32
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IsVisio32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 WindowHandle32
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WindowHandle32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 InstanceHandle
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "InstanceHandle");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 InstanceHandle32
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InstanceHandle32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVUIObject BuiltInMenus
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "BuiltInMenus");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fIgnored">Int16 fIgnored</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVUIObject get_BuiltInToolbars(Int16 fIgnored)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "BuiltInToolbars", NetOffice.VisioApi.IVUIObject.LateBindingApiWrapperType, fIgnored);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_BuiltInToolbars
		/// </summary>
		/// <param name="fIgnored">Int16 fIgnored</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_BuiltInToolbars")]
		public NetOffice.VisioApi.IVUIObject BuiltInToolbars(Int16 fIgnored)
		{
			return get_BuiltInToolbars(fIgnored);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVUIObject CustomMenus
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "CustomMenus");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string CustomMenusFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CustomMenusFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CustomMenusFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVUIObject CustomToolbars
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "CustomToolbars");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string CustomToolbarsFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CustomToolbarsFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CustomToolbarsFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string AddonPaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AddonPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddonPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string DrawingPaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DrawingPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawingPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FilterPaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FilterPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string HelpPaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HelpPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string StartupPaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StartupPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartupPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string StencilPaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StencilPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StencilPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string TemplatePaths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TemplatePaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TemplatePaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string UserName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PromptForSummary
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PromptForSummary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PromptForSummary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVAddons Addons
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVAddons>(this, "Addons");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ProfileName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProfileName");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="eventSeqNum">Int32 eventSeqNum</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_EventInfo(Int32 eventSeqNum)
		{
			return Factory.ExecuteStringPropertyGet(this, "EventInfo", eventSeqNum);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_EventInfo
		/// </summary>
		/// <param name="eventSeqNum">Int32 eventSeqNum</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_EventInfo")]
		public string EventInfo(Int32 eventSeqNum)
		{
			return get_EventInfo(eventSeqNum);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Active
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Active");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 DeferRecalc
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DeferRecalc");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeferRecalc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AlertResponse
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "AlertResponse");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlertResponse", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowProgress
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowProgress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowProgress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object Vbe
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Vbe");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ShowMenus
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowMenus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowMenus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ToolbarStyle
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ToolbarStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ToolbarStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowStatusBar
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowStatusBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowStatusBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 EventsEnabled
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "EventsEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EventsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 TraceFlags
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TraceFlags");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TraceFlags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowToolbar
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowToolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool LiveDynamics
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LiveDynamics");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LiveDynamics", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool AutoLayout
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoLayout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string CommandLine
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandLine");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool IsUndoingOrRedoing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsUndoingOrRedoing");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 CurrentScope
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurrentScope");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nCmdID">Int32 nCmdID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsInScope(Int32 nCmdID)
		{
			return Factory.ExecuteBoolPropertyGet(this, "IsInScope", nCmdID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_IsInScope
		/// </summary>
		/// <param name="nCmdID">Int32 nCmdID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_IsInScope")]
		public bool IsInScope(Int32 nCmdID)
		{
			return get_IsInScope(nCmdID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object old_Addins
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "old_Addins");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ProductName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProductName");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool UndoEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UndoEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UndoEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowChanges
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowChanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 TypelibMajorVersion
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "TypelibMajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 TypelibMinorVersion
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "TypelibMinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AutoRecoverInterval
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "AutoRecoverInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoRecoverInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool InhibitSelectChange
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InhibitSelectChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InhibitSelectChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string ActivePrinter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActivePrinter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ActivePrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public String[] AvailablePrinters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "AvailablePrinters", paramsArray);
				return (String[])returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object CommandBars
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CommandBars");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Build
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object COMAddIns
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "COMAddIns");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultPageUnits
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultPageUnits");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultPageUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object DefaultTextUnits
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultTextUnits");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultTextUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object DefaultAngleUnits
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultAngleUnits");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultAngleUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object DefaultDurationUnits
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultDurationUnits");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultDurationUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 FullBuild
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FullBuild");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool VBAEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VBAEnabled");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisZoomBehavior DefaultZoomBehavior
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisZoomBehavior>(this, "DefaultZoomBehavior");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultZoomBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public stdole.Font DialogFont
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DialogFont", paramsArray);
                return returnItem as stdole.Font;
            }
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 LanguageHelp
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LanguageHelp");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow Window
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "Window");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object ConnectorToolDataObject
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ConnectorToolDataObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplicationSettings Settings
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplicationSettings>(this, "Settings");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object SaveAsWebObject
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "SaveAsWebObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object MsoDebugOptions
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "MsoDebugOptions");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public string MyShapesPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MyShapesPath");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MyShapesPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		public object DefaultRectangleDataObject
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DefaultRectangleDataObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool DataFeaturesEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataFeaturesEnabled");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		public object LanguageSettings
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "LanguageSettings");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		public object Assistance
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Assistance");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool DeferRelationshipRecalc
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DeferRelationshipRecalc");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeferRelationshipRecalc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisEdition CurrentEdition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisEdition>(this, "CurrentEdition");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int64 InstanceHandle64
		{
			get
			{
				return Factory.ExecuteInt64PropertyGet(this, "InstanceHandle64");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Redo()
		{
			 Factory.ExecuteMethod(this, "Redo");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="menusObject">NetOffice.VisioApi.IVUIObject menusObject</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetCustomMenus(NetOffice.VisioApi.IVUIObject menusObject)
		{
			 Factory.ExecuteMethod(this, "SetCustomMenus", menusObject);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ClearCustomMenus()
		{
			 Factory.ExecuteMethod(this, "ClearCustomMenus");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="toolbarsObject">NetOffice.VisioApi.IVUIObject toolbarsObject</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetCustomToolbars(NetOffice.VisioApi.IVUIObject toolbarsObject)
		{
			 Factory.ExecuteMethod(this, "SetCustomToolbars", toolbarsObject);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ClearCustomToolbars()
		{
			 Factory.ExecuteMethod(this, "ClearCustomToolbars");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SaveWorkspaceAs(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveWorkspaceAs", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandID">Int16 commandID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void DoCmd(Int16 commandID)
		{
			 Factory.ExecuteMethod(this, "DoCmd", commandID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="stringOrNumber">object stringOrNumber</param>
		/// <param name="unitsIn">object unitsIn</param>
		/// <param name="unitsOut">object unitsOut</param>
		/// <param name="format">string format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormatResult(object stringOrNumber, object unitsIn, object unitsOut, string format)
		{
			return Factory.ExecuteStringMethodGet(this, "FormatResult", stringOrNumber, unitsIn, unitsOut, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="stringOrNumber">object stringOrNumber</param>
		/// <param name="unitsIn">object unitsIn</param>
		/// <param name="unitsOut">object unitsOut</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double ConvertResult(object stringOrNumber, object unitsIn, object unitsOut)
		{
			return Factory.ExecuteDoubleMethodGet(this, "ConvertResult", stringOrNumber, unitsIn, unitsOut);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pathsString">string pathsString</param>
		/// <param name="nameArray">String[] nameArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void EnumDirectories(string pathsString, out String[] nameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			nameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pathsString, (object)nameArray);
			Invoker.Method(this, "EnumDirectories", paramsArray, modifiers);
			nameArray = (String[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PurgeUndo()
		{
			 Factory.ExecuteMethod(this, "PurgeUndo");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="contextString">string contextString</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 QueueMarkerEvent(string contextString)
		{
			return Factory.ExecuteInt32MethodGet(this, "QueueMarkerEvent", contextString);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrUndoScopeName">string bstrUndoScopeName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 BeginUndoScope(string bstrUndoScopeName)
		{
			return Factory.ExecuteInt32MethodGet(this, "BeginUndoScope", bstrUndoScopeName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nScopeID">Int32 nScopeID</param>
		/// <param name="bCommit">bool bCommit</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void EndUndoScope(Int32 nScopeID, bool bCommit)
		{
			 Factory.ExecuteMethod(this, "EndUndoScope", nScopeID, bCommit);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pUndoUnit">object pUndoUnit</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddUndoUnit(object pUndoUnit)
		{
			 Factory.ExecuteMethod(this, "AddUndoUnit", pUndoUnit);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrScopeName">string bstrScopeName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void RenameCurrentScope(string bstrScopeName)
		{
			 Factory.ExecuteMethod(this, "RenameCurrentScope", bstrScopeName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrHelpFileName">string bstrHelpFileName</param>
		/// <param name="command">Int32 command</param>
		/// <param name="data">Int32 data</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void InvokeHelp(string bstrHelpFileName, Int32 command, Int32 data)
		{
			 Factory.ExecuteMethod(this, "InvokeHelp", bstrHelpFileName, command, data);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="uStateID">NetOffice.VisioApi.Enums.VisOnComponentEnterCodes uStateID</param>
		/// <param name="bEnter">bool bEnter</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void OnComponentEnterState(NetOffice.VisioApi.Enums.VisOnComponentEnterCodes uStateID, bool bEnter)
		{
			 Factory.ExecuteMethod(this, "OnComponentEnterState", uStateID, bEnter);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nWhichStatistic">Int32 nWhichStatistic</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object GetUsageStatistic(Int32 nWhichStatistic)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetUsageStatistic", nWhichStatistic);
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format, object langID, object calendarID)
		{
			return Factory.ExecuteStringMethodGet(this, "FormatResultEx", new object[]{ stringOrNumber, unitsIn, unitsOut, format, langID, calendarID });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="stringOrNumber">object stringOrNumber</param>
		/// <param name="unitsIn">object unitsIn</param>
		/// <param name="unitsOut">object unitsOut</param>
		/// <param name="format">string format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format)
		{
			return Factory.ExecuteStringMethodGet(this, "FormatResultEx", stringOrNumber, unitsIn, unitsOut, format);
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format, object langID)
		{
			return Factory.ExecuteStringMethodGet(this, "FormatResultEx", new object[]{ stringOrNumber, unitsIn, unitsOut, format, langID });
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="sourceAddOn">object sourceAddOn</param>
		/// <param name="targetDocument">NetOffice.VisioApi.IVDocument targetDocument</param>
		/// <param name="targetModes">NetOffice.VisioApi.Enums.VisRibbonXModes targetModes</param>
		/// <param name="friendlyName">string friendlyName</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void RegisterRibbonX(object sourceAddOn, NetOffice.VisioApi.IVDocument targetDocument, NetOffice.VisioApi.Enums.VisRibbonXModes targetModes, string friendlyName)
		{
			 Factory.ExecuteMethod(this, "RegisterRibbonX", sourceAddOn, targetDocument, targetModes, friendlyName);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="sourceAddOn">object sourceAddOn</param>
		/// <param name="targetDocument">NetOffice.VisioApi.IVDocument targetDocument</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void UnregisterRibbonX(object sourceAddOn, NetOffice.VisioApi.IVDocument targetDocument)
		{
			 Factory.ExecuteMethod(this, "UnregisterRibbonX", sourceAddOn, targetDocument);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="galleryName">string galleryName</param>
		[SupportByVersion("Visio", 14,15,16)]
		public bool GetPreviewEnabled(string galleryName)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetPreviewEnabled", galleryName);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="galleryName">string galleryName</param>
		/// <param name="onOrOff">bool onOrOff</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetPreviewEnabled(string galleryName, bool onOrOff)
		{
			 Factory.ExecuteMethod(this, "SetPreviewEnabled", galleryName, onOrOff);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="stencilType">NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType</param>
		/// <param name="measurementSystem">NetOffice.VisioApi.Enums.VisMeasurementSystem measurementSystem</param>
		[SupportByVersion("Visio", 14,15,16)]
		public string GetBuiltInStencilFile(NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType, NetOffice.VisioApi.Enums.VisMeasurementSystem measurementSystem)
		{
			return Factory.ExecuteStringMethodGet(this, "GetBuiltInStencilFile", stencilType, measurementSystem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="stencilType">NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType</param>
		[SupportByVersion("Visio", 14,15,16)]
		public string GetCustomStencilFile(NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType)
		{
			return Factory.ExecuteStringMethodGet(this, "GetCustomStencilFile", stencilType);
		}

		#endregion

		#pragma warning restore
	}
}
