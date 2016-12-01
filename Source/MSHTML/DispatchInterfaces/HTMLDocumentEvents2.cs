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
	/// DispatchInterface HTMLDocumentEvents2 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class HTMLDocumentEvents2 : DispHTMLDocument
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
                    _type = typeof(HTMLDocumentEvents2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLDocumentEvents2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLDocumentEvents2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLDocumentEvents2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLDocumentEvents2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLDocumentEvents2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLDocumentEvents2() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLDocumentEvents2(string progId) : base(progId)
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
		public bool onhelp(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onhelp", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onclick(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onclick", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondblclick(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "ondblclick", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onkeydown(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onkeydown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onkeyup(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onkeyup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onkeypress(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onkeypress", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmousedown(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onmousedown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmousemove(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onmousemove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseup(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onmouseup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseout(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onmouseout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseover(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onmouseover", paramsArray);
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
		public bool ondragstart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "ondragstart", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onselectstart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onselectstart", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
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
		public bool oncontextmenu(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "oncontextmenu", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onstop(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onstop", paramsArray);
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
		public void onpropertychange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onpropertychange", paramsArray);
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
		public void onbeforeeditfocus(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onbeforeeditfocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onselectionchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onselectionchange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool oncontrolselect(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "oncontrolselect", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onmousewheel(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onmousewheel", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onfocusin(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onfocusin", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onfocusout(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onfocusout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "onactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			Invoker.Method(this, "ondeactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onbeforeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforedeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pEvtObj);
			object returnItem = Invoker.MethodReturn(this, "onbeforedeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}