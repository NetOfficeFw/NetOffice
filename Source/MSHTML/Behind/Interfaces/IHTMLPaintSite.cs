using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLPaintSite 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLPaintSite : COMObject, NetOffice.MSHTMLApi.IHTMLPaintSite
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLPaintSite);
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
                    _type = typeof(IHTMLPaintSite);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLPaintSite() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InvalidatePainterInfo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InvalidatePainterInfo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="prcInvalid">tagRECT prcInvalid</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InvalidateRect(tagRECT prcInvalid)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InvalidateRect", prcInvalid);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="rgnInvalid">_RemotableHandle rgnInvalid</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InvalidateRegion(_RemotableHandle rgnInvalid)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InvalidateRegion", rgnInvalid);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pDrawInfo">_HTML_PAINT_DRAW_INFO pDrawInfo</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetDrawInfo(Int32 lFlags, out _HTML_PAINT_DRAW_INFO pDrawInfo)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pDrawInfo = new NetOffice.MSHTMLApi._HTML_PAINT_DRAW_INFO();
			object[] paramsArray = Invoker.ValidateParamsArray(lFlags, pDrawInfo);
			object returnItem = Invoker.MethodReturn(this, "GetDrawInfo", paramsArray, modifiers);
			pDrawInfo = (_HTML_PAINT_DRAW_INFO)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ptGlobal">tagPOINT ptGlobal</param>
		/// <param name="pptLocal">tagPOINT pptLocal</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 TransformGlobalToLocal(tagPOINT ptGlobal, out tagPOINT pptLocal)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pptLocal = new NetOffice.MSHTMLApi.tagPOINT();
			object[] paramsArray = Invoker.ValidateParamsArray(ptGlobal, pptLocal);
			object returnItem = Invoker.MethodReturn(this, "TransformGlobalToLocal", paramsArray, modifiers);
			pptLocal = (tagPOINT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ptLocal">tagPOINT ptLocal</param>
		/// <param name="pptGlobal">tagPOINT pptGlobal</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 TransformLocalToGlobal(tagPOINT ptLocal, out tagPOINT pptGlobal)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pptGlobal = new NetOffice.MSHTMLApi.tagPOINT();
			object[] paramsArray = Invoker.ValidateParamsArray(ptLocal, pptGlobal);
			object returnItem = Invoker.MethodReturn(this, "TransformLocalToGlobal", paramsArray, modifiers);
			pptGlobal = (tagPOINT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plCookie">Int32 plCookie</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetHitTestCookie(out Int32 plCookie)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			plCookie = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(plCookie);
			object returnItem = Invoker.MethodReturn(this, "GetHitTestCookie", paramsArray, modifiers);
			plCookie = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

