using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLDocument4 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLDocument4 : IHTMLDocument3, NetOffice.MSHTMLApi.IHTMLDocument4
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLDocument4);
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
                    _type = typeof(IHTMLDocument4);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLDocument4() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onselectionchange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onselectionchange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onselectionchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object namespaces
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "namespaces");			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string media
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "media");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "media", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object oncontrolselect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "oncontrolselect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "oncontrolselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string URLUnencoded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "URLUnencoded");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void focus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool hasFocus()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "hasFocus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="bstrOptions">string bstrOptions</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLDocument2 createDocumentFromUrl(string bstrUrl, string bstrOptions)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "createDocumentFromUrl", typeof(NetOffice.MSHTMLApi.IHTMLDocument2), bstrUrl, bstrOptions);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLEventObj CreateEventObject(object pvarEventObject)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLEventObj>(this, "CreateEventObject", pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLEventObj CreateEventObject()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLEventObj>(this, "CreateEventObject");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		/// <param name="pvarEventObject">optional object pvarEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool FireEvent(string bstrEventName, object pvarEventObject)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName, pvarEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrEventName">string bstrEventName</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool FireEvent(string bstrEventName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "FireEvent", bstrEventName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLRenderStyle createRenderStyle(string v)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLRenderStyle>(this, "createRenderStyle", typeof(NetOffice.MSHTMLApi.IHTMLRenderStyle), v);
		}

		#endregion

		#pragma warning restore
	}
}


