using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface HTMLFrameSiteEvents 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class HTMLFrameSiteEvents : COMObject
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
                    _type = typeof(HTMLFrameSiteEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public HTMLFrameSiteEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onhelp()
		{
			return Factory.ExecuteBoolMethodGet(this, "onhelp");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onclick()
		{
			return Factory.ExecuteBoolMethodGet(this, "onclick");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ondblclick()
		{
			return Factory.ExecuteBoolMethodGet(this, "ondblclick");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onkeypress()
		{
			return Factory.ExecuteBoolMethodGet(this, "onkeypress");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onkeydown()
		{
			 Factory.ExecuteMethod(this, "onkeydown");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onkeyup()
		{
			 Factory.ExecuteMethod(this, "onkeyup");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmouseout()
		{
			 Factory.ExecuteMethod(this, "onmouseout");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmouseover()
		{
			 Factory.ExecuteMethod(this, "onmouseover");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmousemove()
		{
			 Factory.ExecuteMethod(this, "onmousemove");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmousedown()
		{
			 Factory.ExecuteMethod(this, "onmousedown");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmouseup()
		{
			 Factory.ExecuteMethod(this, "onmouseup");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onselectstart()
		{
			return Factory.ExecuteBoolMethodGet(this, "onselectstart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onfilterchange()
		{
			 Factory.ExecuteMethod(this, "onfilterchange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ondragstart()
		{
			return Factory.ExecuteBoolMethodGet(this, "ondragstart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onbeforeupdate()
		{
			return Factory.ExecuteBoolMethodGet(this, "onbeforeupdate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onafterupdate()
		{
			 Factory.ExecuteMethod(this, "onafterupdate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onerrorupdate()
		{
			return Factory.ExecuteBoolMethodGet(this, "onerrorupdate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onrowexit()
		{
			return Factory.ExecuteBoolMethodGet(this, "onrowexit");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onrowenter()
		{
			 Factory.ExecuteMethod(this, "onrowenter");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void ondatasetchanged()
		{
			 Factory.ExecuteMethod(this, "ondatasetchanged");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void ondataavailable()
		{
			 Factory.ExecuteMethod(this, "ondataavailable");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void ondatasetcomplete()
		{
			 Factory.ExecuteMethod(this, "ondatasetcomplete");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onlosecapture()
		{
			 Factory.ExecuteMethod(this, "onlosecapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onpropertychange()
		{
			 Factory.ExecuteMethod(this, "onpropertychange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onscroll()
		{
			 Factory.ExecuteMethod(this, "onscroll");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onfocus()
		{
			 Factory.ExecuteMethod(this, "onfocus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onblur()
		{
			 Factory.ExecuteMethod(this, "onblur");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onresize()
		{
			 Factory.ExecuteMethod(this, "onresize");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ondrag()
		{
			return Factory.ExecuteBoolMethodGet(this, "ondrag");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void ondragend()
		{
			 Factory.ExecuteMethod(this, "ondragend");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ondragenter()
		{
			return Factory.ExecuteBoolMethodGet(this, "ondragenter");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ondragover()
		{
			return Factory.ExecuteBoolMethodGet(this, "ondragover");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void ondragleave()
		{
			 Factory.ExecuteMethod(this, "ondragleave");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool ondrop()
		{
			return Factory.ExecuteBoolMethodGet(this, "ondrop");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onbeforecut()
		{
			return Factory.ExecuteBoolMethodGet(this, "onbeforecut");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool oncut()
		{
			return Factory.ExecuteBoolMethodGet(this, "oncut");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onbeforecopy()
		{
			return Factory.ExecuteBoolMethodGet(this, "onbeforecopy");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool oncopy()
		{
			return Factory.ExecuteBoolMethodGet(this, "oncopy");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onbeforepaste()
		{
			return Factory.ExecuteBoolMethodGet(this, "onbeforepaste");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onpaste()
		{
			return Factory.ExecuteBoolMethodGet(this, "onpaste");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool oncontextmenu()
		{
			return Factory.ExecuteBoolMethodGet(this, "oncontextmenu");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onrowsdelete()
		{
			 Factory.ExecuteMethod(this, "onrowsdelete");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onrowsinserted()
		{
			 Factory.ExecuteMethod(this, "onrowsinserted");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void oncellchange()
		{
			 Factory.ExecuteMethod(this, "oncellchange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onreadystatechange()
		{
			 Factory.ExecuteMethod(this, "onreadystatechange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onbeforeeditfocus()
		{
			 Factory.ExecuteMethod(this, "onbeforeeditfocus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onlayoutcomplete()
		{
			 Factory.ExecuteMethod(this, "onlayoutcomplete");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onpage()
		{
			 Factory.ExecuteMethod(this, "onpage");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onbeforedeactivate()
		{
			return Factory.ExecuteBoolMethodGet(this, "onbeforedeactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onbeforeactivate()
		{
			return Factory.ExecuteBoolMethodGet(this, "onbeforeactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmove()
		{
			 Factory.ExecuteMethod(this, "onmove");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool oncontrolselect()
		{
			return Factory.ExecuteBoolMethodGet(this, "oncontrolselect");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onmovestart()
		{
			return Factory.ExecuteBoolMethodGet(this, "onmovestart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmoveend()
		{
			 Factory.ExecuteMethod(this, "onmoveend");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onresizestart()
		{
			return Factory.ExecuteBoolMethodGet(this, "onresizestart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onresizeend()
		{
			 Factory.ExecuteMethod(this, "onresizeend");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmouseenter()
		{
			 Factory.ExecuteMethod(this, "onmouseenter");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onmouseleave()
		{
			 Factory.ExecuteMethod(this, "onmouseleave");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool onmousewheel()
		{
			return Factory.ExecuteBoolMethodGet(this, "onmousewheel");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onactivate()
		{
			 Factory.ExecuteMethod(this, "onactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void ondeactivate()
		{
			 Factory.ExecuteMethod(this, "ondeactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onfocusin()
		{
			 Factory.ExecuteMethod(this, "onfocusin");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onfocusout()
		{
			 Factory.ExecuteMethod(this, "onfocusout");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void onload()
		{
			 Factory.ExecuteMethod(this, "onload");
		}

		#endregion

		#pragma warning restore
	}
}
