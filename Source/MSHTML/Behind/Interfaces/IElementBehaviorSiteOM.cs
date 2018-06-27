using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IElementBehaviorSiteOM 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IElementBehaviorSiteOM : COMObject, NetOffice.MSHTMLApi.IElementBehaviorSiteOM
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IElementBehaviorSiteOM);
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
                    _type = typeof(IElementBehaviorSiteOM);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IElementBehaviorSiteOM() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchEvent">string pchEvent</param>
		/// <param name="lFlags">Int32 lFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RegisterEvent(string pchEvent, Int32 lFlags)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RegisterEvent", pchEvent, lFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchEvent">string pchEvent</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetEventCookie(string pchEvent)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetEventCookie", pchEvent);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lCookie">Int32 lCookie</param>
		/// <param name="pEventObject">NetOffice.MSHTMLApi.IHTMLEventObj pEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 FireEvent(Int32 lCookie, NetOffice.MSHTMLApi.IHTMLEventObj pEventObject)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FireEvent", lCookie, pEventObject);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLEventObj CreateEventObject()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLEventObj>(this, "CreateEventObject");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchName">string pchName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RegisterName(string pchName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RegisterName", pchName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchUrn">string pchUrn</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RegisterUrn(string pchUrn)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RegisterUrn", pchUrn);
		}

		#endregion

		#pragma warning restore
	}
}

