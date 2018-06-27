using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLDocumentEvents2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class HTMLDocumentEvents2 : DispHTMLDocument, NetOffice.MSHTMLApi.HTMLDocumentEvents2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.HTMLDocumentEvents2);
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
                    _type = typeof(HTMLDocumentEvents2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLDocumentEvents2() : base()
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
		public virtual bool onkeypress(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onkeypress", pEvtObj);
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
		public virtual void onmousemove(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onmousemove", pEvtObj);
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
		public virtual void onreadystatechange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onreadystatechange", pEvtObj);
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
		public virtual bool ondragstart(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ondragstart", pEvtObj);
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
		public virtual bool onerrorupdate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onerrorupdate", pEvtObj);
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
		public virtual bool onstop(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onstop", pEvtObj);
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
		public virtual void onpropertychange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onpropertychange", pEvtObj);
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
		public virtual void onbeforeeditfocus(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onbeforeeditfocus", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onselectionchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onselectionchange", pEvtObj);
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
		public virtual bool onmousewheel(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onmousewheel", pEvtObj);
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
		public virtual bool onbeforeactivate(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onbeforeactivate", pEvtObj);
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

		#endregion

		#pragma warning restore
	}
}

