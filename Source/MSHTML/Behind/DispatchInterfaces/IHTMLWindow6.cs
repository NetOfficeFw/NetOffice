using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLWindow6 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLWindow6 : IHTMLWindow5, NetOffice.MSHTMLApi.IHTMLWindow6
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLWindow6);
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
                    _type = typeof(IHTMLWindow6);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLWindow6() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object XDomainRequest
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "XDomainRequest");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "XDomainRequest", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStorage sessionStorage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStorage>(this, "sessionStorage", typeof(NetOffice.MSHTMLApi.IHTMLStorage));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLStorage localStorage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLStorage>(this, "localStorage", typeof(NetOffice.MSHTMLApi.IHTMLStorage));
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onhashchange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onhashchange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onhashchange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 maxConnectionsPerServer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "maxConnectionsPerServer");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onmessage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onmessage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onmessage", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="msg">string msg</param>
		/// <param name="targetOrigin">optional object targetOrigin</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void postMessage(string msg, object targetOrigin)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "postMessage", msg, targetOrigin);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="msg">string msg</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void postMessage(string msg)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "postMessage", msg);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHTML">string bstrHTML</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string toStaticHTML(string bstrHTML)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "toStaticHTML", bstrHTML);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrProfilerMarkName">string bstrProfilerMarkName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void msWriteProfilerMark(string bstrProfilerMarkName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "msWriteProfilerMark", bstrProfilerMarkName);
		}

		#endregion

		#pragma warning restore
	}
}


