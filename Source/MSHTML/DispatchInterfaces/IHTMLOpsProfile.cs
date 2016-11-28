using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// DispatchInterface IHTMLOpsProfile 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IHTMLOpsProfile : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IHTMLOpsProfile);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLOpsProfile(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool addRequest(string name, object reserved)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, reserved);
			object returnItem = Invoker.MethodReturn(this, "addRequest", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool addRequest(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "addRequest", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void clearRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "clearRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain, object path, object expire, object reserved)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain, path, expire, reserved);
			Invoker.Method(this, "doRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doRequest(object usage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage);
			Invoker.Method(this, "doRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doRequest(object usage, object fname)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname);
			Invoker.Method(this, "doRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain);
			Invoker.Method(this, "doRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain, object path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain, path);
			Invoker.Method(this, "doRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain, object path, object expire)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain, path, expire);
			Invoker.Method(this, "doRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public string getAttribute(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "getAttribute", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		/// <param name="prefs">optional object prefs</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool setAttribute(string name, string value, object prefs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, value, prefs);
			object returnItem = Invoker.MethodReturn(this, "setAttribute", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool setAttribute(string name, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, value);
			object returnItem = Invoker.MethodReturn(this, "setAttribute", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool commitChanges()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "commitChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool addReadRequest(string name, object reserved)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, reserved);
			object returnItem = Invoker.MethodReturn(this, "addReadRequest", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool addReadRequest(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "addReadRequest", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain, object path, object expire, object reserved)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain, path, expire, reserved);
			Invoker.Method(this, "doReadRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doReadRequest(object usage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage);
			Invoker.Method(this, "doReadRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doReadRequest(object usage, object fname)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname);
			Invoker.Method(this, "doReadRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain);
			Invoker.Method(this, "doReadRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain, object path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain, path);
			Invoker.Method(this, "doReadRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain, object path, object expire)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(usage, fname, domain, path, expire);
			Invoker.Method(this, "doReadRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool doWriteRequest()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "doWriteRequest", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}