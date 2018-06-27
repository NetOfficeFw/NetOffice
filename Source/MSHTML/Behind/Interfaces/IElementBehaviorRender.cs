using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IElementBehaviorRender 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IElementBehaviorRender : COMObject, NetOffice.MSHTMLApi.IElementBehaviorRender
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IElementBehaviorRender);
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
                    _type = typeof(IElementBehaviorRender);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IElementBehaviorRender() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hdc">_RemotableHandle hdc</param>
		/// <param name="lLayer">Int32 lLayer</param>
		/// <param name="pRect">tagRECT pRect</param>
		/// <param name="pReserved">object pReserved</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Draw(_RemotableHandle hdc, Int32 lLayer, tagRECT pRect, object pReserved)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Draw", hdc, lLayer, pRect, pReserved);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetRenderInfo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetRenderInfo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="pReserved">object pReserved</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 HitTestPoint(tagPOINT pPoint, object pReserved)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HitTestPoint", pPoint, pReserved);
		}

		#endregion

		#pragma warning restore
	}
}

