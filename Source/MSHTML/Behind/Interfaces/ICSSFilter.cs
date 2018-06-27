using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface ICSSFilter 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ICSSFilter : COMObject, NetOffice.MSHTMLApi.ICSSFilter
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
                    _contractType = typeof(NetOffice.MSHTMLApi.ICSSFilter);
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
                    _type = typeof(ICSSFilter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ICSSFilter() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSink">NetOffice.MSHTMLApi.ICSSFilterSite pSink</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetSite(NetOffice.MSHTMLApi.ICSSFilterSite pSink)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetSite", pSink);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 OnAmbientPropertyChange(Int32 dispid)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnAmbientPropertyChange", dispid);
		}

		#endregion

		#pragma warning restore
	}
}

