using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLDocumentEvents3 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class HTMLDocumentEvents3 : HTMLDocumentEvents2, NetOffice.MSHTMLApi.HTMLDocumentEvents3
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
                    _contractType = typeof(NetOffice.MSHTMLApi.HTMLDocumentEvents3);
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
                    _type = typeof(HTMLDocumentEvents3);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLDocumentEvents3() : base()
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
		public virtual void onstorage(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onstorage", pEvtObj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void onstoragecommit(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "onstoragecommit", pEvtObj);
		}

		#endregion

		#pragma warning restore
	}
}

