using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLXMLHttpRequest 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLXMLHttpRequest : COMObject
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
                    _type = typeof(DispHTMLXMLHttpRequest);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispHTMLXMLHttpRequest(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLXMLHttpRequest(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLXMLHttpRequest(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLXMLHttpRequest(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLXMLHttpRequest(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLXMLHttpRequest(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLXMLHttpRequest() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLXMLHttpRequest(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 readyState
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object responseBody
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "responseBody");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string responseText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "responseText");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object responseXML
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "responseXML");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 status
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "status");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string statusText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "statusText");
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
		public Int32 timeout
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "timeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "timeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object ontimeout
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ontimeout");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ontimeout", value);
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
		[SupportByVersion("MSHTML", 4)]
		public void abort()
		{
			 Factory.ExecuteMethod(this, "abort");
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
		public void open(string bstrMethod, string bstrUrl, object varAsync, object varUser, object varPassword)
		{
			 Factory.ExecuteMethod(this, "open", new object[]{ bstrMethod, bstrUrl, varAsync, varUser, varPassword });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void open(string bstrMethod, string bstrUrl, object varAsync)
		{
			 Factory.ExecuteMethod(this, "open", bstrMethod, bstrUrl, varAsync);
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
		public void open(string bstrMethod, string bstrUrl, object varAsync, object varUser)
		{
			 Factory.ExecuteMethod(this, "open", bstrMethod, bstrUrl, varAsync, varUser);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varBody">optional object varBody</param>
		[SupportByVersion("MSHTML", 4)]
		public void send(object varBody)
		{
			 Factory.ExecuteMethod(this, "send", varBody);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void send()
		{
			 Factory.ExecuteMethod(this, "send");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string getAllResponseHeaders()
		{
			return Factory.ExecuteStringMethodGet(this, "getAllResponseHeaders");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHeader">string bstrHeader</param>
		[SupportByVersion("MSHTML", 4)]
		public string getResponseHeader(string bstrHeader)
		{
			return Factory.ExecuteStringMethodGet(this, "getResponseHeader", bstrHeader);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHeader">string bstrHeader</param>
		/// <param name="bstrValue">string bstrValue</param>
		[SupportByVersion("MSHTML", 4)]
		public void setRequestHeader(string bstrHeader, string bstrValue)
		{
			 Factory.ExecuteMethod(this, "setRequestHeader", bstrHeader, bstrValue);
		}

		#endregion

		#pragma warning restore
	}
}
