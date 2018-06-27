using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispCEventObj 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispCEventObj : COMObject, NetOffice.MSHTMLApi.DispCEventObj
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispCEventObj);
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
                    _type = typeof(DispCEventObj);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispCEventObj() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object returnValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "returnValue");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "returnValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool cancelBubble
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "cancelBubble");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "cancelBubble", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 keyCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "keyCode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "keyCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string propertyName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "propertyName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "propertyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLBookmarkCollection bookmarks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLBookmarkCollection>(this, "bookmarks", typeof(NetOffice.MSHTMLApi.IHTMLBookmarkCollection));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "bookmarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object recordset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "recordset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string dataFld
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "dataFld");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "dataFld", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElementCollection boundElements
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElementCollection>(this, "boundElements");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "boundElements", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool repeat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "repeat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "repeat", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string srcUrn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "srcUrn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "srcUrn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement srcElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "srcElement");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "srcElement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool altKey
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "altKey");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "altKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ctrlKey
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ctrlKey");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ctrlKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool shiftKey
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "shiftKey");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "shiftKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement fromElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "fromElement");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "fromElement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement toElement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "toElement");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "toElement", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 button
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "button");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "button", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "type", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string qualifier
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "qualifier");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "qualifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 reason
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "reason");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "reason", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 x
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "x");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "x", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 y
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "y");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "y", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 clientX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "clientX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "clientX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 clientY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "clientY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "clientY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 offsetX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "offsetX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "offsetX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 offsetY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "offsetY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "offsetY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 screenX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "screenX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "screenX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 screenY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "screenY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "screenY", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object srcFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "srcFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "srcFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLDataTransfer dataTransfer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDataTransfer>(this, "dataTransfer", typeof(NetOffice.MSHTMLApi.IHTMLDataTransfer));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool contentOverflow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "contentOverflow");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool shiftLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "shiftLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "shiftLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool altLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "altLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "altLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ctrlLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ctrlLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ctrlLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 imeCompositionChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "imeCompositionChange");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 imeNotifyCommand
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "imeNotifyCommand");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 imeNotifyData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "imeNotifyData");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 imeRequest
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "imeRequest");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 imeRequestData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "imeRequestData");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 keyboardLayout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "keyboardLayout");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 behaviorCookie
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "behaviorCookie");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 behaviorPart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "behaviorPart");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string nextPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "nextPage");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 wheelDelta
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "wheelDelta");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string url
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "url");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "url", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string data
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "data");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "data", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object source
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "source");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string origin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "origin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "origin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool issession
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "issession");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "issession", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object constructor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "constructor");
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
		public virtual void setAttribute(string strAttributeName, object attributeValue, object lFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="attributeValue">object attributeValue</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void setAttribute(string strAttributeName, object attributeValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setAttribute", strAttributeName, attributeValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object getAttribute(string strAttributeName, object lFlags)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object getAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getAttribute", strAttributeName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		/// <param name="lFlags">optional Int32 lFlags = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeAttribute(string strAttributeName, object lFlags)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="strAttributeName">string strAttributeName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool removeAttribute(string strAttributeName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "removeAttribute", strAttributeName);
		}

		#endregion

		#pragma warning restore
	}
}


