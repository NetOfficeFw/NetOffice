using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLMarqueeElementEvents2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class HTMLMarqueeElementEvents2 : DispHTMLMarqueeElement, NetOffice.MSHTMLApi.HTMLMarqueeElementEvents2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.HTMLMarqueeElementEvents2);
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
                    _type = typeof(HTMLMarqueeElementEvents2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLMarqueeElementEvents2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onhelp(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onhelp", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onclick(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onclick", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondblclick(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondblclick", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onkeypress(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onkeypress", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onkeydown(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onkeydown", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onkeyup(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onkeyup", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseout(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseout", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseover(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseover", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmousemove(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmousemove", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmousedown(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmousedown", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseup(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseup", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onselectstart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onselectstart", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfilterchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfilterchange", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondragstart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragstart", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforeupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforeupdate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onafterupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onafterupdate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onerrorupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onerrorupdate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onrowexit(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onrowexit", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onrowenter(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onrowenter", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondatasetchanged(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondatasetchanged", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondataavailable(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondataavailable", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondatasetcomplete(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondatasetcomplete", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onlosecapture(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onlosecapture", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onpropertychange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onpropertychange", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onscroll(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onscroll", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfocus(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfocus", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onblur(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onblur", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onresize(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onresize", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondrag(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondrag", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondragend(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondragend", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondragenter(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragenter", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondragover(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragover", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondragleave(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondragleave", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool ondrop(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondrop", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforecut(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforecut", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncut(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncut", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforecopy(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforecopy", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncopy(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncopy", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforepaste(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforepaste", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onpaste(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onpaste", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncontextmenu(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncontextmenu", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onrowsdelete(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onrowsdelete", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onrowsinserted(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onrowsinserted", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void oncellchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "oncellchange", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onreadystatechange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onreadystatechange", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onlayoutcomplete(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onlayoutcomplete", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onpage(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onpage", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseenter(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseenter", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmouseleave(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmouseleave", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onactivate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void ondeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ondeactivate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforedeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforedeactivate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onbeforeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforeactivate", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfocusin(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfocusin", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfocusout(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfocusout", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmove(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmove", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool oncontrolselect(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "oncontrolselect", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onmovestart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onmovestart", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onmoveend(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmoveend", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onresizestart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onresizestart", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onresizeend(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onresizeend", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool onmousewheel(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onmousewheel", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onchange", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onselect(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onselect", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onbounce(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onbounce", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onfinish(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onfinish", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onstart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onstart", pEvtObj);
		}

		#endregion

		#pragma warning restore
	}
}

