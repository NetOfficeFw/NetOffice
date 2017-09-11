using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// IVWindow
	/// </summary>
	[SyntaxBypass]
 	public class IVWindow_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVWindow_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVWindow_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="reviewerID">optional Int32 reviewerID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_ReviewerMarkupVisible(object reviewerID)
		{
			return Factory.ExecuteBoolPropertyGet(this, "ReviewerMarkupVisible", reviewerID);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="reviewerID">optional Int32 reviewerID</param>
        /// <param name="value">optional bool value</param>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ReviewerMarkupVisible(object reviewerID, bool value)
		{
			Factory.ExecutePropertySet(this, "ReviewerMarkupVisible", reviewerID, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ReviewerMarkupVisible
		/// </summary>
		/// <param name="reviewerID">optional Int32 reviewerID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ReviewerMarkupVisible")]
		public bool ReviewerMarkupVisible(object reviewerID)
		{
			return get_ReviewerMarkupVisible(reviewerID);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface IVWindow 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVWindow : IVWindow_
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
                    _type = typeof(IVWindow);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVWindow(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVWindow(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindow(string progId) : base(progId)
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
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Type
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVPage PageAsObj
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "PageAsObj");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string PageFromName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PageFromName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageFromName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double Zoom
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Zoom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Zoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection Selection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "Selection");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Selection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Index
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 SubType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "SubType");
			}
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
		public Int32 WindowHandle32
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WindowHandle32");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowRulers
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowRulers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowRulers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowGrid
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowGuides
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowConnectPoints
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowConnectPoints");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowConnectPoints", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowPageBreaks
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowPageBreaks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPageBreaks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object Page
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Page");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Page", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object Master
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Master");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ShowScrollBars
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ShowScrollBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowScrollBars", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 WindowState
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WindowState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ViewFit
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ViewFit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool IsEditingText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsEditingText");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool IsEditingOLE
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsEditingOLE");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindows Parent
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindows>(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMasterShortcut MasterShortcut
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMasterShortcut>(this, "MasterShortcut");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow ParentWindow
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ParentWindow");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string MergeID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MergeID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MergeID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string MergeClass
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MergeClass");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MergeClass", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 MergePosition
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MergePosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MergePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool AllowEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double PageTabWidth
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "PageTabWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageTabWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowPageTabs
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPageTabs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPageTabs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool InPlace
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InPlace");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string MergeCaption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MergeCaption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MergeCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public stdole.Picture Icon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Icon", paramsArray);
                return returnItem as stdole.Picture;
            }
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Icon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Shape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVCell SelectedCell
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "SelectedCell");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 BackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 BackgroundColorGradient
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackgroundColorGradient");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundColorGradient", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowPageOutline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPageOutline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPageOutline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ScrollLock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ScrollLock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollLock", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ZoomLock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ZoomLock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ZoomLock", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisZoomBehavior ZoomBehavior
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisZoomBehavior>(this, "ZoomBehavior");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ZoomBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object[] SelectedMasters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SelectedMasters", paramsArray);
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this,(object[])returnItem, false);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVCharacters SelectedText
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCharacters>(this, "SelectedText");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SelectedText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ReviewerMarkupVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReviewerMarkupVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReviewerMarkupVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDataRecordset SelectedDataRecordset
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataRecordset>(this, "SelectedDataRecordset");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SelectedDataRecordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32 SelectedDataRowID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelectedDataRowID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelectedDataRowID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection SelectionForDragCopy
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SelectionForDragCopy");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVValidationIssue SelectedValidationIssue
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVValidationIssue>(this, "SelectedValidationIssue");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SelectedValidationIssue", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void DeselectAll()
		{
			 Factory.ExecuteMethod(this, "DeselectAll");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="selectAction">Int16 selectAction</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction)
		{
			 Factory.ExecuteMethod(this, "Select", sheetObject, selectAction);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Duplicate()
		{
			 Factory.ExecuteMethod(this, "Duplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Group()
		{
			 Factory.ExecuteMethod(this, "Group");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Union()
		{
			 Factory.ExecuteMethod(this, "Union");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Combine()
		{
			 Factory.ExecuteMethod(this, "Combine");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Fragment()
		{
			 Factory.ExecuteMethod(this, "Fragment");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddToGroup()
		{
			 Factory.ExecuteMethod(this, "AddToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void RemoveFromGroup()
		{
			 Factory.ExecuteMethod(this, "RemoveFromGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Intersect()
		{
			 Factory.ExecuteMethod(this, "Intersect");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Subtract()
		{
			 Factory.ExecuteMethod(this, "Subtract");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Trim()
		{
			 Factory.ExecuteMethod(this, "Trim");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Join()
		{
			 Factory.ExecuteMethod(this, "Join");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nameArray">String[] nameArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void DockedStencils(out String[] nameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			nameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)nameArray);
			Invoker.Method(this, "DockedStencils", paramsArray, modifiers);
			nameArray = (String[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nxFlags">Int32 nxFlags</param>
		/// <param name="nyFlags">Int32 nyFlags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Scroll(Int32 nxFlags, Int32 nyFlags)
		{
			 Factory.ExecuteMethod(this, "Scroll", nxFlags, nyFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ScrollViewTo(Double x, Double y)
		{
			 Factory.ExecuteMethod(this, "ScrollViewTo", x, y);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pdLeft">Double pdLeft</param>
		/// <param name="pdTop">Double pdTop</param>
		/// <param name="pdWidth">Double pdWidth</param>
		/// <param name="pdHeight">Double pdHeight</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void GetViewRect(out Double pdLeft, out Double pdTop, out Double pdWidth, out Double pdHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pdLeft = 0;
			pdTop = 0;
			pdWidth = 0;
			pdHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pdLeft, pdTop, pdWidth, pdHeight);
			Invoker.Method(this, "GetViewRect", paramsArray, modifiers);
			pdLeft = (Double)paramsArray[0];
			pdTop = (Double)paramsArray[1];
			pdWidth = (Double)paramsArray[2];
			pdHeight = (Double)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dLeft">Double dLeft</param>
		/// <param name="dTop">Double dTop</param>
		/// <param name="dWidth">Double dWidth</param>
		/// <param name="dHeight">Double dHeight</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetViewRect(Double dLeft, Double dTop, Double dWidth, Double dHeight)
		{
			 Factory.ExecuteMethod(this, "SetViewRect", dLeft, dTop, dWidth, dHeight);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pnLeft">Int32 pnLeft</param>
		/// <param name="pnTop">Int32 pnTop</param>
		/// <param name="pnWidth">Int32 pnWidth</param>
		/// <param name="pnHeight">Int32 pnHeight</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void GetWindowRect(out Int32 pnLeft, out Int32 pnTop, out Int32 pnWidth, out Int32 pnHeight)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pnLeft = 0;
			pnTop = 0;
			pnWidth = 0;
			pnHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pnLeft, pnTop, pnWidth, pnHeight);
			Invoker.Method(this, "GetWindowRect", paramsArray, modifiers);
			pnLeft = (Int32)paramsArray[0];
			pnTop = (Int32)paramsArray[1];
			pnWidth = (Int32)paramsArray[2];
			pnHeight = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nLeft">Int32 nLeft</param>
		/// <param name="nTop">Int32 nTop</param>
		/// <param name="nWidth">Int32 nWidth</param>
		/// <param name="nHeight">Int32 nHeight</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetWindowRect(Int32 nLeft, Int32 nTop, Int32 nWidth, Int32 nHeight)
		{
			 Factory.ExecuteMethod(this, "SetWindowRect", nLeft, nTop, nWidth, nHeight);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow NewWindow()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "NewWindow");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisCenterViewFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void CenterViewOnShape(NetOffice.VisioApi.IVShape sheetObject, NetOffice.VisioApi.Enums.VisCenterViewFlags flags)
		{
			 Factory.ExecuteMethod(this, "CenterViewOnShape", sheetObject, flags);
		}

		#endregion

		#pragma warning restore
	}
}
