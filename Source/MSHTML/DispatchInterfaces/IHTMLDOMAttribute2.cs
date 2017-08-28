using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLDOMAttribute2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLDOMAttribute2 : IHTMLDOMAttribute
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
                    _type = typeof(IHTMLDOMAttribute2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLDOMAttribute2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLDOMAttribute2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMAttribute2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMAttribute2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMAttribute2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMAttribute2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMAttribute2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMAttribute2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "name");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "value");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool expando
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "expando");
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

		#endregion

		#region Methods

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
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMNode appendChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMNode>(this, "appendChild", newChild);
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
		/// <param name="fDeep">bool fDeep</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute cloneNode(bool fDeep)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "cloneNode", fDeep);
		}

		#endregion

		#pragma warning restore
	}
}
