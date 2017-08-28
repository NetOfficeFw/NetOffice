using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLElement4 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLElement4 : IHTMLElement3
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
                    _type = typeof(IHTMLElement4);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLElement4(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement4(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement4(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement4(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement4(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement4() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLElement4(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void normalize()
		{
			 Factory.ExecuteMethod(this, "normalize");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute getAttributeNode(string bstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "getAttributeNode", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute setAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "setAttributeNode", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, pattr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pattr">NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute removeAttributeNode(NetOffice.MSHTMLApi.IHTMLDOMAttribute pattr)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "removeAttributeNode", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, pattr);
		}

		#endregion

		#pragma warning restore
	}
}



