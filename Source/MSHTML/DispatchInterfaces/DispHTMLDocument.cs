using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLDocument 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLDocument : COMObject
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
                    _type = typeof(DispHTMLDocument);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispHTMLDocument(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLDocument(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLDocument(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLDocument(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLDocument(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLDocument(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLDocument() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLDocument(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Script
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Script");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection all
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "all");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement body
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "body");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement activeElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "activeElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection images
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "images");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection applets
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "applets");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection links
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "links");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection forms
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "forms");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection anchors
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "anchors");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "title", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection scripts
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "scripts");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string designMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "designMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "designMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLSelectionObject selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLSelectionObject>(this, "selection", NetOffice.MSHTMLApi.IHTMLSelectionObject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string readyState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLFramesCollection2 frames
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFramesCollection2>(this, "frames");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection embeds
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "embeds");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection plugins
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "plugins");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object alinkColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "alinkColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "alinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object bgColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "bgColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "bgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object fgColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "fgColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "fgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object linkColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "linkColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "linkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object vlinkColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "vlinkColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "vlinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string referrer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "referrer");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLLocation location
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLLocation>(this, "location", NetOffice.MSHTMLApi.IHTMLLocation.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string lastModified
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "lastModified");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string url
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "url");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "url", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string domain
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "domain");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "domain", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string cookie
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "cookie");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "cookie", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool expando
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "expando");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "expando", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string charset
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "charset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "charset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string defaultCharset
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "defaultCharset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "defaultCharset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string mimeType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "mimeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileSize
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileSize");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileCreatedDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileCreatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileModifiedDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileModifiedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string fileUpdatedDate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "fileUpdatedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string security
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "security");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string protocol
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "protocol");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string nameProp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "nameProp");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onhelp
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onhelp");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onhelp", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onclick
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onclick");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onclick", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondblclick
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondblclick");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondblclick", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onkeyup
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onkeyup");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onkeyup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onkeydown
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onkeydown");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onkeydown", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onkeypress
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onkeypress");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onkeypress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseup
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseup");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmousedown
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmousedown");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmousedown", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmousemove
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmousemove");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmousemove", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseout
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseout");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmouseover
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmouseover");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmouseover", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onreadystatechange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onreadystatechange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onreadystatechange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onafterupdate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onafterupdate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onafterupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowexit
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowexit");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowexit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowenter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowenter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondragstart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondragstart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondragstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onselectstart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onselectstart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onselectstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 parentWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "parentWindow", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection styleSheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection>(this, "styleSheets", NetOffice.MSHTMLApi.IHTMLStyleSheetsCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeupdate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeupdate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onerrorupdate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onerrorupdate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onerrorupdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement documentElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "documentElement");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string uniqueID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "uniqueID");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowsdelete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowsdelete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowsdelete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onrowsinserted
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onrowsinserted");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onrowsinserted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncellchange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncellchange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncellchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondatasetchanged
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondatasetchanged");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondatasetchanged", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondataavailable
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondataavailable");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondataavailable", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondatasetcomplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondatasetcomplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondatasetcomplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onpropertychange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onpropertychange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onpropertychange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dir
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dir");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dir", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncontextmenu
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncontextmenu");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncontextmenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onstop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onstop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onstop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSHTMLApi.IHTMLDocument2 parentDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "parentDocument", NetOffice.MSHTMLApi.IHTMLDocument2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool enableDownload
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "enableDownload");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "enableDownload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string baseUrl
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "baseUrl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "baseUrl", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool inheritStyleSheets
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "inheritStyleSheets");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "inheritStyleSheets", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeeditfocus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeeditfocus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeeditfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onselectionchange
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onselectionchange");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onselectionchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object namespaces
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "namespaces");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string media
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "media");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "media", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object oncontrolselect
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "oncontrolselect");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "oncontrolselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string URLUnencoded
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "URLUnencoded");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onmousewheel
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onmousewheel");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onmousewheel", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode doctype
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "doctype");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMImplementation implementation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMImplementation>(this, "implementation", NetOffice.MSHTMLApi.IHTMLDOMImplementation.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocusin
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocusin");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocusin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocusout
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocusout");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocusout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ondeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ondeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ondeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforedeactivate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforedeactivate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforedeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string compatMode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "compatMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDocumentCompatibleInfoCollection compatible
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDocumentCompatibleInfoCollection>(this, "compatible", NetOffice.MSHTMLApi.IHTMLDocumentCompatibleInfoCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object documentMode
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "documentMode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onstorage
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onstorage");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onstorage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onstoragecommit
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onstoragecommit");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onstoragecommit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 nodeType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "nodeType");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode parentNode
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "parentNode");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object childNodes
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "childNodes");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object attributes
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "attributes");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string nodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "nodeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object nodeValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "nodeValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "nodeValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode firstChild
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "firstChild");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode lastChild
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "lastChild");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode previousSibling
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "previousSibling");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode nextSibling
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "nextSibling");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object ownerDocument
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ownerDocument");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object constructor
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "constructor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psarray">optional object[] psarray</param>
		[SupportByVersion("MSHTML", 4)]
		public void write(object[] psarray)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)psarray);
            Invoker.Method(this, "write", paramsArray);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void write()
		{
			 Factory.ExecuteMethod(this, "write");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psarray">optional object[] psarray</param>
		[SupportByVersion("MSHTML", 4)]
		public void writeln(object[] psarray)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)psarray);
            Invoker.Method(this, "writeln", paramsArray);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void writeln()
		{
			 Factory.ExecuteMethod(this, "writeln");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		/// <param name="name">optional object name</param>
		/// <param name="features">optional object features</param>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("MSHTML", 4)]
		public object open(object url, object name, object features, object replace)
		{
			return Factory.ExecuteVariantMethodGet(this, "open", url, name, features, replace);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object open()
		{
			return Factory.ExecuteVariantMethodGet(this, "open");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object open(object url)
		{
			return Factory.ExecuteVariantMethodGet(this, "open", url);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object open(object url, object name)
		{
			return Factory.ExecuteVariantMethodGet(this, "open", url, name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = text/html</param>
		/// <param name="name">optional object name</param>
		/// <param name="features">optional object features</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object open(object url, object name, object features)
		{
			return Factory.ExecuteVariantMethodGet(this, "open", url, name, features);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void close()
		{
			 Factory.ExecuteMethod(this, "close");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void clear()
		{
			 Factory.ExecuteMethod(this, "clear");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandSupported(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandSupported", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandEnabled(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandEnabled", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandState(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandState", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandIndeterm(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandIndeterm", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public string queryCommandText(string cmdID)
		{
			return Factory.ExecuteStringMethodGet(this, "queryCommandText", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public object queryCommandValue(string cmdID)
		{
			return Factory.ExecuteVariantMethodGet(this, "queryCommandValue", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("MSHTML", 4)]
		public bool execCommand(string cmdID, object showUI, object value)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool execCommand(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommand", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool execCommand(string cmdID, object showUI)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool execCommandShowHelp(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommandShowHelp", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eTag">string eTag</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement createElement(string eTag)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "createElement", eTag);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement elementFromPoint(Int32 x, Int32 y)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "elementFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string toString()
		{
			return Factory.ExecuteStringMethodGet(this, "toString");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHref">optional string bstrHref = </param>
		/// <param name="lIndex">optional Int32 lIndex = -1</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLStyleSheet createStyleSheet(object bstrHref, object lIndex)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "createStyleSheet", bstrHref, lIndex);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLStyleSheet createStyleSheet()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "createStyleSheet");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHref">optional string bstrHref = </param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLStyleSheet createStyleSheet(object bstrHref)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLStyleSheet>(this, "createStyleSheet", bstrHref);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void releaseCapture()
		{
			 Factory.ExecuteMethod(this, "releaseCapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fForce">optional bool fForce = false</param>
		[SupportByVersion("MSHTML", 4)]
		public void recalc(object fForce)
		{
			 Factory.ExecuteMethod(this, "recalc", fForce);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void recalc()
		{
			 Factory.ExecuteMethod(this, "recalc");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode createTextNode(string text)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "createTextNode", text);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_event">string event</param>
		/// <param name="pdisp">object pdisp</param>
		[SupportByVersion("MSHTML", 4)]
		public bool attachEvent(string _event, object pdisp)
		{
			return Factory.ExecuteBoolMethodGet(this, "attachEvent", _event, pdisp);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_event">string event</param>
		/// <param name="pdisp">object pdisp</param>
		[SupportByVersion("MSHTML", 4)]
		public void detachEvent(string _event, object pdisp)
		{
			 Factory.ExecuteMethod(this, "detachEvent", _event, pdisp);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDocument2 createDocumentFragment()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "createDocumentFragment", NetOffice.MSHTMLApi.IHTMLDocument2.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection getElementsByName(string v)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "getElementsByName", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement getElementById(string v)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "getElementById", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection getElementsByTagName(string v)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "getElementsByTagName", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void focus()
		{
			 Factory.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hasFocus()
		{
			return Factory.ExecuteBoolMethodGet(this, "hasFocus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="bstrOptions">string bstrOptions</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDocument2 createDocumentFromUrl(string bstrUrl, string bstrOptions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "createDocumentFromUrl", NetOffice.MSHTMLApi.IHTMLDocument2.LateBindingApiWrapperType, bstrUrl, bstrOptions);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLEventObj CreateEventObject(object pvarEventObject)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLEventObj>(this, "CreateEventObject", pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLEventObj CreateEventObject()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLEventObj>(this, "CreateEventObject");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		public bool FireEvent(string bstrEventName, object pvarEventObject)
		{
			return Factory.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName, pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool FireEvent(string bstrEventName)
		{
			return Factory.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLRenderStyle createRenderStyle(string v)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLRenderStyle>(this, "createRenderStyle", NetOffice.MSHTMLApi.IHTMLRenderStyle.LateBindingApiWrapperType, v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrattrName">string bstrattrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute createAttribute(string bstrattrName)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "createAttribute", bstrattrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrdata">string bstrdata</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode createComment(string bstrdata)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "createComment", bstrdata);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrId">string bstrId</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement2 ie8_getElementById(string bstrId)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement2>(this, "ie8_getElementById", bstrId);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void updateSettings()
		{
			 Factory.ExecuteMethod(this, "updateSettings");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool hasChildNodes()
		{
			return Factory.ExecuteBoolMethodGet(this, "hasChildNodes");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="refChild">optional object refChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, object refChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "insertBefore", newChild, refChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "insertBefore", newChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode removeChild(NetOffice.MSHTMLApi.IHTMLDOMNode oldChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeChild", oldChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode replaceChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, NetOffice.MSHTMLApi.IHTMLDOMNode oldChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "replaceChild", newChild, oldChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">bool fDeep</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode cloneNode(bool fDeep)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "cloneNode", fDeep);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">optional bool fDeep = false</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode removeNode(object fDeep)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeNode", fDeep);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMNode removeNode()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "removeNode");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="otherNode">NetOffice.MSHTMLApi.IHTMLDOMNode otherNode</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode swapNode(NetOffice.MSHTMLApi.IHTMLDOMNode otherNode)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "swapNode", otherNode);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="replacement">NetOffice.MSHTMLApi.IHTMLDOMNode replacement</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode replaceNode(NetOffice.MSHTMLApi.IHTMLDOMNode replacement)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "replaceNode", replacement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode appendChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "appendChild", newChild);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement querySelector(string v)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "querySelector", v);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection querySelectorAll(string v)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection>(this, "querySelectorAll", NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection.LateBindingApiWrapperType, v);
		}

		#endregion

		#pragma warning restore
	}
}
