using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IElementBehavior 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IElementBehavior : COMObject, NetOffice.MSHTMLApi.IElementBehavior
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IElementBehavior);
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
                    _type = typeof(IElementBehavior);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IElementBehavior() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pBehaviorSite">NetOffice.MSHTMLApi.IElementBehaviorSite pBehaviorSite</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Init(NetOffice.MSHTMLApi.IElementBehaviorSite pBehaviorSite)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Init", pBehaviorSite);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lEvent">Int32 lEvent</param>
		/// <param name="pVar">object pVar</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Notify(Int32 lEvent, object pVar)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Notify", lEvent, pVar);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Detach()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Detach");
		}

		#endregion

		#pragma warning restore
	}
}

