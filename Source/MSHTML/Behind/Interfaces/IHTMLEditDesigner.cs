using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLEditDesigner 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLEditDesigner : COMObject, NetOffice.MSHTMLApi.IHTMLEditDesigner
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLEditDesigner);
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
                    _type = typeof(IHTMLEditDesigner);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLEditDesigner() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 PreHandleEvent(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PreHandleEvent", inEvtDispId, pIEventObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 PostHandleEvent(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostHandleEvent", inEvtDispId, pIEventObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 TranslateAccelerator(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "TranslateAccelerator", inEvtDispId, pIEventObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 PostEditorEventNotify(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostEditorEventNotify", inEvtDispId, pIEventObj);
		}

		#endregion

		#pragma warning restore
	}
}

