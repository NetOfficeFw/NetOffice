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
	/// DispatchInterface HTMLObjectElementEvents2 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class HTMLObjectElementEvents2 : DispHTMLObjectElement
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
                    _type = typeof(HTMLObjectElementEvents2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLObjectElementEvents2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLObjectElementEvents2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLObjectElementEvents2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLObjectElementEvents2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLObjectElementEvents2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLObjectElementEvents2() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLObjectElementEvents2(string progId) : base(progId)
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
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforeupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onbeforeupdate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onafterupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onafterupdate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onerrorupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onerrorupdate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onrowexit(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onrowexit", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onrowenter(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onrowenter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondatasetchanged(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "ondatasetchanged", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondataavailable(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "ondataavailable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondatasetcomplete(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "ondatasetcomplete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onerror(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onerror", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onrowsdelete(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onrowsdelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onrowsinserted(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onrowsinserted", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void oncellchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "oncellchange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onreadystatechange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onreadystatechange", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}