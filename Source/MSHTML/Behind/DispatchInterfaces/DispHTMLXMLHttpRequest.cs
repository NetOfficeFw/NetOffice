using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLXMLHttpRequest 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLXMLHttpRequest : COMObject, NetOffice.MSHTMLApi.DispHTMLXMLHttpRequest
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLXMLHttpRequest);
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
                    _type = typeof(DispHTMLXMLHttpRequest);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLXMLHttpRequest() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 readyState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object responseBody
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "responseBody");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string responseText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "responseText");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public virtual object responseXML
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "responseXML");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 status
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "status");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string statusText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "statusText");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onreadystatechange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onreadystatechange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onreadystatechange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 timeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "timeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "timeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object ontimeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ontimeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ontimeout", value);
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
		[SupportByVersion("MSHTML", 4)]
		public virtual void abort()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "abort");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		/// <param name="varUser">optional object varUser</param>
		/// <param name="varPassword">optional object varPassword</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void open(string bstrMethod, string bstrUrl, object varAsync, object varUser, object varPassword)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "open", new object[]{ bstrMethod, bstrUrl, varAsync, varUser, varPassword });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void open(string bstrMethod, string bstrUrl, object varAsync)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "open", bstrMethod, bstrUrl, varAsync);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		/// <param name="varUser">optional object varUser</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void open(string bstrMethod, string bstrUrl, object varAsync, object varUser)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "open", bstrMethod, bstrUrl, varAsync, varUser);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varBody">optional object varBody</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void send(object varBody)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "send", varBody);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void send()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "send");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string getAllResponseHeaders()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getAllResponseHeaders");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHeader">string bstrHeader</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string getResponseHeader(string bstrHeader)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getResponseHeader", bstrHeader);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHeader">string bstrHeader</param>
		/// <param name="bstrValue">string bstrValue</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void setRequestHeader(string bstrHeader, string bstrValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setRequestHeader", bstrHeader, bstrValue);
		}

		#endregion

		#pragma warning restore
	}
}

