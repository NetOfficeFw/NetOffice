using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface IVInvisibleApp 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IVInvisibleApp : COMObject, NetOffice.VisioApi.IVInvisibleApp
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
                    _contractType = typeof(NetOffice.VisioApi.IVInvisibleApp);
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
                    _type = typeof(IVInvisibleApp);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVInvisibleApp() : base()
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
		public virtual NetOffice.VisioApi.IVDocument ActiveDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "ActiveDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVPage ActivePage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ActivePage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVWindow ActiveWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ActiveWindow");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocuments Documents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocuments>(this, "Documents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 OnDataChangeDelay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "OnDataChangeDelay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDataChangeDelay", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ProcessID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcessID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ScreenUpdating
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ScreenUpdating");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 WindowHandle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WindowHandle");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVWindows Windows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindows>(this, "Windows");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Language
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Language");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 IsVisio16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IsVisio16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 IsVisio32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IsVisio32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 WindowHandle32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WindowHandle32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 InstanceHandle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "InstanceHandle");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 InstanceHandle32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "InstanceHandle32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVUIObject BuiltInMenus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "BuiltInMenus");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fIgnored">Int16 fIgnored</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVUIObject get_BuiltInToolbars(Int16 fIgnored)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "BuiltInToolbars", typeof(NetOffice.VisioApi.IVUIObject), fIgnored);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_BuiltInToolbars
		/// </summary>
		/// <param name="fIgnored">Int16 fIgnored</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_BuiltInToolbars")]
		public virtual NetOffice.VisioApi.IVUIObject BuiltInToolbars(Int16 fIgnored)
		{
			return get_BuiltInToolbars(fIgnored);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVUIObject CustomMenus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "CustomMenus");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string CustomMenusFile
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CustomMenusFile");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CustomMenusFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVUIObject CustomToolbars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "CustomToolbars");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string CustomToolbarsFile
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CustomToolbarsFile");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CustomToolbarsFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string AddonPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddonPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddonPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string DrawingPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DrawingPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawingPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string FilterPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FilterPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FilterPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string HelpPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HelpPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string StartupPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StartupPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartupPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string StencilPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StencilPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StencilPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string TemplatePaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TemplatePaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TemplatePaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string UserName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 PromptForSummary
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PromptForSummary");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PromptForSummary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVAddons Addons
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVAddons>(this, "Addons");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ProfileName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProfileName");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="eventSeqNum">Int32 eventSeqNum</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_EventInfo(Int32 eventSeqNum)
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EventInfo", eventSeqNum);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_EventInfo
		/// </summary>
		/// <param name="eventSeqNum">Int32 eventSeqNum</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_EventInfo")]
		public virtual string EventInfo(Int32 eventSeqNum)
		{
			return get_EventInfo(eventSeqNum);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 PersistsEvents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Active
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Active");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 DeferRecalc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DeferRecalc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeferRecalc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 AlertResponse
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "AlertResponse");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlertResponse", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ShowProgress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowProgress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowProgress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object Vbe
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Vbe");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 ShowMenus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowMenus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowMenus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 ToolbarStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ToolbarStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ToolbarStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ShowStatusBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowStatusBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowStatusBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 EventsEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "EventsEnabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EventsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 TraceFlags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TraceFlags");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TraceFlags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ShowToolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShowToolbar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool LiveDynamics
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LiveDynamics");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LiveDynamics", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool AutoLayout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoLayout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string CommandLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandLine");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool IsUndoingOrRedoing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsUndoingOrRedoing");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 CurrentScope
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentScope");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nCmdID">Int32 nCmdID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool get_IsInScope(Int32 nCmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsInScope", nCmdID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_IsInScope
		/// </summary>
		/// <param name="nCmdID">Int32 nCmdID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_IsInScope")]
		public virtual bool IsInScope(Int32 nCmdID)
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
		public virtual object old_Addins
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "old_Addins");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ProductName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProductName");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool UndoEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UndoEnabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UndoEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ShowChanges
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowChanges");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 TypelibMajorVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "TypelibMajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 TypelibMinorVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "TypelibMinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 AutoRecoverInterval
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "AutoRecoverInterval");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoRecoverInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool InhibitSelectChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InhibitSelectChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InhibitSelectChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string ActivePrinter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ActivePrinter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ActivePrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual String[] AvailablePrinters
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
		public virtual object CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CommandBars");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Build
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object COMAddIns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "COMAddIns");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object DefaultPageUnits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultPageUnits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultPageUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual object DefaultTextUnits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultTextUnits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultTextUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual object DefaultAngleUnits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultAngleUnits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultAngleUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual object DefaultDurationUnits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultDurationUnits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultDurationUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 FullBuild
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FullBuild");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool VBAEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "VBAEnabled");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisZoomBehavior DefaultZoomBehavior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisZoomBehavior>(this, "DefaultZoomBehavior");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultZoomBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public virtual stdole.Font DialogFont
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
		public virtual Int32 LanguageHelp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LanguageHelp");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVWindow Window
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "Window");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object ConnectorToolDataObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ConnectorToolDataObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVApplicationSettings Settings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplicationSettings>(this, "Settings");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object SaveAsWebObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "SaveAsWebObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object MsoDebugOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "MsoDebugOptions");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual string MyShapesPath
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MyShapesPath");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MyShapesPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		public virtual object DefaultRectangleDataObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DefaultRectangleDataObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual bool DataFeaturesEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DataFeaturesEnabled");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		public virtual object LanguageSettings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "LanguageSettings");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16), ProxyResult]
		public virtual object Assistance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Assistance");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool DeferRelationshipRecalc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DeferRelationshipRecalc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeferRelationshipRecalc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisEdition CurrentEdition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisEdition>(this, "CurrentEdition");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int64 InstanceHandle64
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt64PropertyGet(this, "InstanceHandle64");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Quit()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Redo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Redo");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Undo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="menusObject">NetOffice.VisioApi.IVUIObject menusObject</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SetCustomMenus(NetOffice.VisioApi.IVUIObject menusObject)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCustomMenus", menusObject);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ClearCustomMenus()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearCustomMenus");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="toolbarsObject">NetOffice.VisioApi.IVUIObject toolbarsObject</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SetCustomToolbars(NetOffice.VisioApi.IVUIObject toolbarsObject)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCustomToolbars", toolbarsObject);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ClearCustomToolbars()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearCustomToolbars");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SaveWorkspaceAs(string fileName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveWorkspaceAs", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandID">Int16 commandID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 DoCmd(Int16 commandID)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DoCmd", commandID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="stringOrNumber">object stringOrNumber</param>
		/// <param name="unitsIn">object unitsIn</param>
		/// <param name="unitsOut">object unitsOut</param>
		/// <param name="format">string format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FormatResult(object stringOrNumber, object unitsIn, object unitsOut, string format)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "FormatResult", stringOrNumber, unitsIn, unitsOut, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="stringOrNumber">object stringOrNumber</param>
		/// <param name="unitsIn">object unitsIn</param>
		/// <param name="unitsOut">object unitsOut</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Double ConvertResult(object stringOrNumber, object unitsIn, object unitsOut)
		{
			return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ConvertResult", stringOrNumber, unitsIn, unitsOut);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pathsString">string pathsString</param>
		/// <param name="nameArray">String[] nameArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 EnumDirectories(string pathsString, out String[] nameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			nameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pathsString, (object)nameArray);
			object returnItem = Invoker.MethodReturn(this, "EnumDirectories", paramsArray);
			nameArray = (String[])paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 PurgeUndo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PurgeUndo");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="contextString">string contextString</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 QueueMarkerEvent(string contextString)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "QueueMarkerEvent", contextString);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrUndoScopeName">string bstrUndoScopeName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 BeginUndoScope(string bstrUndoScopeName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeginUndoScope", bstrUndoScopeName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nScopeID">Int32 nScopeID</param>
		/// <param name="bCommit">bool bCommit</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 EndUndoScope(Int32 nScopeID, bool bCommit)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndUndoScope", nScopeID, bCommit);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pUndoUnit">object pUndoUnit</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 AddUndoUnit(object pUndoUnit)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AddUndoUnit", pUndoUnit);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrScopeName">string bstrScopeName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 RenameCurrentScope(string bstrScopeName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RenameCurrentScope", bstrScopeName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrHelpFileName">string bstrHelpFileName</param>
		/// <param name="command">Int32 command</param>
		/// <param name="data">Int32 data</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 InvokeHelp(string bstrHelpFileName, Int32 command, Int32 data)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InvokeHelp", bstrHelpFileName, command, data);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="uStateID">NetOffice.VisioApi.Enums.VisOnComponentEnterCodes uStateID</param>
		/// <param name="bEnter">bool bEnter</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 OnComponentEnterState(NetOffice.VisioApi.Enums.VisOnComponentEnterCodes uStateID, bool bEnter)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnComponentEnterState", uStateID, bEnter);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nWhichStatistic">Int32 nWhichStatistic</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual object GetUsageStatistic(Int32 nWhichStatistic)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetUsageStatistic", nWhichStatistic);
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
		public virtual string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format, object langID, object calendarID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "FormatResultEx", new object[]{ stringOrNumber, unitsIn, unitsOut, format, langID, calendarID });
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
		public virtual string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "FormatResultEx", stringOrNumber, unitsIn, unitsOut, format);
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
		public virtual string FormatResultEx(object stringOrNumber, object unitsIn, object unitsOut, string format, object langID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "FormatResultEx", new object[]{ stringOrNumber, unitsIn, unitsOut, format, langID });
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="sourceAddOn">object sourceAddOn</param>
		/// <param name="targetDocument">NetOffice.VisioApi.IVDocument targetDocument</param>
		/// <param name="targetModes">NetOffice.VisioApi.Enums.VisRibbonXModes targetModes</param>
		/// <param name="friendlyName">string friendlyName</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 RegisterRibbonX(object sourceAddOn, NetOffice.VisioApi.IVDocument targetDocument, NetOffice.VisioApi.Enums.VisRibbonXModes targetModes, string friendlyName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RegisterRibbonX", sourceAddOn, targetDocument, targetModes, friendlyName);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="sourceAddOn">object sourceAddOn</param>
		/// <param name="targetDocument">NetOffice.VisioApi.IVDocument targetDocument</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 UnregisterRibbonX(object sourceAddOn, NetOffice.VisioApi.IVDocument targetDocument)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UnregisterRibbonX", sourceAddOn, targetDocument);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="galleryName">string galleryName</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool GetPreviewEnabled(string galleryName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetPreviewEnabled", galleryName);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="galleryName">string galleryName</param>
		/// <param name="onOrOff">bool onOrOff</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 SetPreviewEnabled(string galleryName, bool onOrOff)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetPreviewEnabled", galleryName, onOrOff);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="stencilType">NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType</param>
		/// <param name="measurementSystem">NetOffice.VisioApi.Enums.VisMeasurementSystem measurementSystem</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual string GetBuiltInStencilFile(NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType, NetOffice.VisioApi.Enums.VisMeasurementSystem measurementSystem)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBuiltInStencilFile", stencilType, measurementSystem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="stencilType">NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual string GetCustomStencilFile(NetOffice.VisioApi.Enums.VisBuiltInStencilTypes stencilType)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCustomStencilFile", stencilType);
		}

		#endregion

		#pragma warning restore
	}
}


