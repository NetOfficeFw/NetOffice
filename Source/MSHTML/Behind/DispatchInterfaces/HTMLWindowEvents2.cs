using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLWindowEvents2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class HTMLWindowEvents2 : DispHTMLWindow2, NetOffice.MSHTMLApi.HTMLWindowEvents2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.HTMLWindowEvents2);
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
                    _type = typeof(HTMLWindowEvents2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLWindowEvents2() : base()
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
		public virtual void onload(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onload", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onunload(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onunload", pEvtObj);
		}

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
		/// <param name="description">string description</param>
		/// <param name="url">string url</param>
		/// <param name="line">Int32 line</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onerror(string description, string url, Int32 line)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onerror", description, url, line);
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
		public virtual void onscroll(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onscroll", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onbeforeunload(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onbeforeunload", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onbeforeprint(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onbeforeprint", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onafterprint(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onafterprint", pEvtObj);
		}

		#endregion

		#pragma warning restore
	}
}

