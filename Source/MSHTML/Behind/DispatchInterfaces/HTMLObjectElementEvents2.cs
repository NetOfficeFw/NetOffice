using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLObjectElementEvents2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class HTMLObjectElementEvents2 : DispHTMLObjectElement, NetOffice.MSHTMLApi.HTMLObjectElementEvents2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.HTMLObjectElementEvents2);
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
                    _type = typeof(HTMLObjectElementEvents2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLObjectElementEvents2() : base()
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
		public virtual bool onerror(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "onerror", pEvtObj);
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

		#endregion

		#pragma warning restore
	}
}

