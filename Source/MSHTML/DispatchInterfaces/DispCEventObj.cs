using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispCEventObj 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispCEventObj : COMObject
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
                    _type = typeof(DispCEventObj);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispCEventObj(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispCEventObj(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispCEventObj(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispCEventObj(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispCEventObj(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispCEventObj(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispCEventObj() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispCEventObj(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object returnValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "returnValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "returnValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool cancelBubble
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "cancelBubble");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "cancelBubble", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 keyCode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "keyCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "keyCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string propertyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "propertyName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "propertyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLBookmarkCollection bookmarks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLBookmarkCollection>(this, "bookmarks", NetOffice.MSHTMLApi.IHTMLBookmarkCollection.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "bookmarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object recordset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "recordset");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string dataFld
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "dataFld");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "dataFld", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElementCollection boundElements
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "boundElements");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "boundElements", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool repeat
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "repeat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "repeat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string srcUrn
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "srcUrn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "srcUrn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement srcElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "srcElement");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "srcElement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool altKey
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "altKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "altKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ctrlKey
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ctrlKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ctrlKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool shiftKey
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "shiftKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "shiftKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement fromElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "fromElement");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "fromElement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement toElement
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "toElement");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "toElement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 button
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "button");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "button", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string type
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "type");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "type", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string qualifier
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "qualifier");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "qualifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 reason
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "reason");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "reason", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 x
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "x");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "x", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 y
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "y");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "y", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 clientX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "clientX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "clientX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 clientY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "clientY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "clientY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 offsetX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "offsetX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "offsetX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 offsetY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "offsetY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "offsetY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 screenX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "screenX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "screenX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 screenY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "screenY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "screenY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object srcFilter
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "srcFilter");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "srcFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDataTransfer dataTransfer
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDataTransfer>(this, "dataTransfer", NetOffice.MSHTMLApi.IHTMLDataTransfer.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool contentOverflow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "contentOverflow");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool shiftLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "shiftLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "shiftLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool altLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "altLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "altLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ctrlLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ctrlLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ctrlLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 imeCompositionChange
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "imeCompositionChange");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 imeNotifyCommand
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "imeNotifyCommand");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 imeNotifyData
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "imeNotifyData");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 imeRequest
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "imeRequest");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 imeRequestData
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "imeRequestData");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 keyboardLayout
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "keyboardLayout");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 behaviorCookie
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "behaviorCookie");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 behaviorPart
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "behaviorPart");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string nextPage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "nextPage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 wheelDelta
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "wheelDelta");
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
		public string data
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "data");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "data", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object source
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "source");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string origin
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "origin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "origin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool issession
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "issession");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "issession", value);
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
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		/// <param name="lFlags">optional Int32 lFlags = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public void setAttribute(string strAttributeName, object attributeValue, object lFlags)
		{
			 Factory.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void setAttribute(string strAttributeName, object attributeValue)
		{
			 Factory.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		public object getAttribute(string strAttributeName, object lFlags)
		{
			return Factory.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object getAttribute(string strAttributeName)
		{
			return Factory.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public bool removeAttribute(string strAttributeName, object lFlags)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool removeAttribute(string strAttributeName)
		{
			return Factory.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName);
		}

		#endregion

		#pragma warning restore
	}
}
