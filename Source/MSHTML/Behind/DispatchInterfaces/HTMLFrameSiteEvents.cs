using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLFrameSiteEvents 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class HTMLFrameSiteEvents : COMObject, NetOffice.MSHTMLApi.HTMLFrameSiteEvents
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
                    _contractType = typeof(NetOffice.MSHTMLApi.HTMLFrameSiteEvents);
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
                    _type = typeof(HTMLFrameSiteEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLFrameSiteEvents() : base()
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
		public virtual bool onhelp()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onhelp");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onclick()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onclick");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondblclick()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondblclick");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onkeypress()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onkeypress");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onkeydown()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onkeydown");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onkeyup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onkeyup");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseout");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseover()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseover");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmousemove()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmousemove");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmousedown()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmousedown");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseup");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onselectstart()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onselectstart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfilterchange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfilterchange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondragstart()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragstart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforeupdate()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforeupdate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onafterupdate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onafterupdate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onerrorupdate()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onerrorupdate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onrowexit()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onrowexit");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onrowenter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onrowenter");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondatasetchanged()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondatasetchanged");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondataavailable()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondataavailable");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondatasetcomplete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondatasetcomplete");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onlosecapture()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onlosecapture");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onpropertychange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onpropertychange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onscroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onscroll");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfocus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfocus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onblur()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onblur");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onresize()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onresize");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondrag()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondrag");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondragend()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondragend");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondragenter()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragenter");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondragover()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragover");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondragleave()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondragleave");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondrop()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondrop");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforecut()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforecut");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncut()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncut");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforecopy()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforecopy");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncopy()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncopy");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforepaste()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforepaste");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onpaste()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onpaste");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncontextmenu()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncontextmenu");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onrowsdelete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onrowsdelete");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onrowsinserted()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onrowsinserted");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void oncellchange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "oncellchange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onreadystatechange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onreadystatechange");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onbeforeeditfocus()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onbeforeeditfocus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onlayoutcomplete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onlayoutcomplete");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onpage()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onpage");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforedeactivate()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforedeactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforeactivate()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforeactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmove()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmove");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncontrolselect()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncontrolselect");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onmovestart()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onmovestart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmoveend()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmoveend");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onresizestart()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onresizestart");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onresizeend()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onresizeend");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseenter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseenter");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseleave()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseleave");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onmousewheel()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onmousewheel");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onactivate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondeactivate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondeactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfocusin()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfocusin");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfocusout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfocusout");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onload()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onload");
		}

		#endregion

		#pragma warning restore
	}
}

