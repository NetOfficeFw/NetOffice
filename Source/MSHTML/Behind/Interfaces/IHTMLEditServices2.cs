using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLEditServices2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLEditServices2 : IHTMLEditServices, NetOffice.MSHTMLApi.IHTMLEditServices2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLEditServices2);
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
                    _type = typeof(IHTMLEditServices2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLEditServices2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IDisplayPointer pIStartAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToSelectionAnchorEx(NetOffice.MSHTMLApi.IDisplayPointer pIStartAnchor)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToSelectionAnchorEx", pIStartAnchor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IDisplayPointer pIEndAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToSelectionEndEx(NetOffice.MSHTMLApi.IDisplayPointer pIEndAnchor)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToSelectionEndEx", pIEndAnchor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fReCompute">Int32 fReCompute</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 FreezeVirtualCaretPos(Int32 fReCompute)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FreezeVirtualCaretPos", fReCompute);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fReset">Int32 fReset</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 UnFreezeVirtualCaretPos(Int32 fReset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UnFreezeVirtualCaretPos", fReset);
		}

		#endregion

		#pragma warning restore
	}
}

