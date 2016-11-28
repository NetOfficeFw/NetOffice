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
	/// DispatchInterface HTMLFrameSiteEvents 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class HTMLFrameSiteEvents : COMObject
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
                    _type = typeof(HTMLFrameSiteEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLFrameSiteEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameSiteEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameSiteEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameSiteEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameSiteEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameSiteEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLFrameSiteEvents(string progId) : base(progId)
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
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onhelp()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onhelp", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onclick()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onclick", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondblclick()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ondblclick", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onkeypress()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onkeypress", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onkeydown()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onkeydown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onkeyup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onkeyup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseout()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmouseout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseover()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmouseover", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmousemove()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmousemove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmousedown()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmousedown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmouseup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onselectstart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onselectstart", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onfilterchange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onfilterchange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondragstart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ondragstart", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforeupdate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onbeforeupdate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onafterupdate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onafterupdate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onerrorupdate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onerrorupdate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onrowexit()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onrowexit", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onrowenter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onrowenter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondatasetchanged()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ondatasetchanged", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondataavailable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ondataavailable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondatasetcomplete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ondatasetcomplete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onlosecapture()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onlosecapture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onpropertychange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onpropertychange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onscroll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onscroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onfocus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onfocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onblur()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onblur", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onresize()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onresize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondrag()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ondrag", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondragend()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ondragend", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondragenter()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ondragenter", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondragover()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ondragover", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondragleave()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ondragleave", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool ondrop()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ondrop", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforecut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onbeforecut", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool oncut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "oncut", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforecopy()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onbeforecopy", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool oncopy()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "oncopy", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforepaste()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onbeforepaste", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onpaste()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onpaste", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool oncontextmenu()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "oncontextmenu", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onrowsdelete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onrowsdelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onrowsinserted()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onrowsinserted", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void oncellchange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "oncellchange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onreadystatechange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onreadystatechange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onbeforeeditfocus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onbeforeeditfocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onlayoutcomplete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onlayoutcomplete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onpage()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onpage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforedeactivate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onbeforedeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onbeforeactivate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onbeforeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmove()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool oncontrolselect()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "oncontrolselect", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onmovestart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onmovestart", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmoveend()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmoveend", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onresizestart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onresizestart", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onresizeend()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onresizeend", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseenter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmouseenter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onmouseleave()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onmouseleave", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool onmousewheel()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "onmousewheel", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onactivate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void ondeactivate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ondeactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onfocusin()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onfocusin", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onfocusout()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onfocusout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void onload()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "onload", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}