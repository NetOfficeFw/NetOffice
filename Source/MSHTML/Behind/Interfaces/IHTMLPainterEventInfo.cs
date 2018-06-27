using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLPainterEventInfo 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLPainterEventInfo : COMObject, NetOffice.MSHTMLApi.IHTMLPainterEventInfo
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLPainterEventInfo);
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
                    _type = typeof(IHTMLPainterEventInfo);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLPainterEventInfo() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plEventInfoFlags">Int32 plEventInfoFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetEventInfoFlags(out Int32 plEventInfoFlags)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			plEventInfoFlags = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(plEventInfoFlags);
			object returnItem = Invoker.MethodReturn(this, "GetEventInfoFlags", paramsArray, modifiers);
			plEventInfoFlags = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetEventTarget(NetOffice.MSHTMLApi.IHTMLElement ppElement)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetEventTarget", ppElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lPartID">Int32 lPartID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCursor(Int32 lPartID)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCursor", lPartID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lPartID">Int32 lPartID</param>
		/// <param name="pbstrPart">string pbstrPart</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 StringFromPartID(Int32 lPartID, out string pbstrPart)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrPart = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(lPartID, pbstrPart);
			object returnItem = Invoker.MethodReturn(this, "StringFromPartID", paramsArray, modifiers);
			pbstrPart = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

