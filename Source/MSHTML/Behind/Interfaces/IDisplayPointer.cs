using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IDisplayPointer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IDisplayPointer : COMObject, NetOffice.MSHTMLApi.IDisplayPointer
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IDisplayPointer);
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
                    _type = typeof(IDisplayPointer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDisplayPointer() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ptPoint">tagPOINT ptPoint</param>
		/// <param name="eCoordSystem">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eCoordSystem</param>
		/// <param name="pElementContext">NetOffice.MSHTMLApi.IHTMLElement pElementContext</param>
		/// <param name="dwHitTestOptions">Int32 dwHitTestOptions</param>
		/// <param name="pdwHitTestResults">Int32 pdwHitTestResults</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 moveToPoint(tagPOINT ptPoint, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eCoordSystem, NetOffice.MSHTMLApi.IHTMLElement pElementContext, Int32 dwHitTestOptions, out Int32 pdwHitTestResults)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			pdwHitTestResults = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(ptPoint, eCoordSystem, pElementContext, dwHitTestOptions, pdwHitTestResults);
			object returnItem = Invoker.MethodReturn(this, "moveToPoint", paramsArray, modifiers);
			pdwHitTestResults = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eMoveUnit">NetOffice.MSHTMLApi.Enums._DISPLAY_MOVEUNIT eMoveUnit</param>
		/// <param name="lXPos">Int32 lXPos</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveUnit(NetOffice.MSHTMLApi.Enums._DISPLAY_MOVEUNIT eMoveUnit, Int32 lXPos)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUnit", eMoveUnit, lXPos);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pMarkupPointer">NetOffice.MSHTMLApi.IMarkupPointer pMarkupPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 PositionMarkupPointer(NetOffice.MSHTMLApi.IMarkupPointer pMarkupPointer)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PositionMarkupPointer", pMarkupPointer);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToPointer(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToPointer", pDispPointer);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY eGravity</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetPointerGravity(NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY eGravity)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetPointerGravity", eGravity);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY peGravity</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetPointerGravity(out NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY peGravity)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			peGravity = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(peGravity);
			object returnItem = Invoker.MethodReturn(this, "GetPointerGravity", paramsArray, modifiers);
			peGravity = (NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eGravity">NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY eGravity</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetDisplayGravity(NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY eGravity)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetDisplayGravity", eGravity);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peGravity">NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY peGravity</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetDisplayGravity(out NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY peGravity)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			peGravity = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(peGravity);
			object returnItem = Invoker.MethodReturn(this, "GetDisplayGravity", paramsArray, modifiers);
			peGravity = (NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfPositioned">Int32 pfPositioned</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsPositioned(out Int32 pfPositioned)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfPositioned = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfPositioned);
			object returnItem = Invoker.MethodReturn(this, "IsPositioned", paramsArray, modifiers);
			pfPositioned = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Unposition()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Unposition");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="pfIsEqual">Int32 pfIsEqual</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsEqualTo(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, out Int32 pfIsEqual)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfIsEqual = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointer, pfIsEqual);
			object returnItem = Invoker.MethodReturn(this, "IsEqualTo", paramsArray, modifiers);
			pfIsEqual = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="pfIsLeftOf">Int32 pfIsLeftOf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsLeftOf(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, out Int32 pfIsLeftOf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfIsLeftOf = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointer, pfIsLeftOf);
			object returnItem = Invoker.MethodReturn(this, "IsLeftOf", paramsArray, modifiers);
			pfIsLeftOf = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="pfIsRightOf">Int32 pfIsRightOf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsRightOf(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, out Int32 pfIsRightOf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfIsRightOf = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointer, pfIsRightOf);
			object returnItem = Invoker.MethodReturn(this, "IsRightOf", paramsArray, modifiers);
			pfIsRightOf = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfBOL">Int32 pfBOL</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsAtBOL(out Int32 pfBOL)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfBOL = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfBOL);
			object returnItem = Invoker.MethodReturn(this, "IsAtBOL", paramsArray, modifiers);
			pfBOL = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		/// <param name="pDispLineContext">NetOffice.MSHTMLApi.IDisplayPointer pDispLineContext</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToMarkupPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointer, NetOffice.MSHTMLApi.IDisplayPointer pDispLineContext)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToMarkupPointer", pPointer, pDispLineContext);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollIntoView()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "scrollIntoView");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppLineInfo">NetOffice.MSHTMLApi.ILineInfo ppLineInfo</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetLineInfo(out NetOffice.MSHTMLApi.ILineInfo ppLineInfo)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppLineInfo = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppLineInfo);
			object returnItem = Invoker.MethodReturn(this, "GetLineInfo", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppLineInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.ILineInfo>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.ILineInfo));
            else
                ppLineInfo = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppLayoutElement">NetOffice.MSHTMLApi.IHTMLElement ppLayoutElement</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetFlowElement(out NetOffice.MSHTMLApi.IHTMLElement ppLayoutElement)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppLayoutElement = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppLayoutElement);
			object returnItem = Invoker.MethodReturn(this, "GetFlowElement", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppLayoutElement = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppLayoutElement = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pdwBreaks">Int32 pdwBreaks</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 QueryBreaks(out Int32 pdwBreaks)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pdwBreaks = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pdwBreaks);
			object returnItem = Invoker.MethodReturn(this, "QueryBreaks", paramsArray, modifiers);
			pdwBreaks = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

