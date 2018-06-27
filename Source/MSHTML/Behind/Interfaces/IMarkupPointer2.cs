using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupPointer2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IMarkupPointer2 : IMarkupPointer, NetOffice.MSHTMLApi.IMarkupPointer2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupPointer2);
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
                    _type = typeof(IMarkupPointer2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupPointer2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfAtBreak">Int32 pfAtBreak</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsAtWordBreak(out Int32 pfAtBreak)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfAtBreak = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfAtBreak);
			object returnItem = Invoker.MethodReturn(this, "IsAtWordBreak", paramsArray, modifiers);
			pfAtBreak = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plMP">Int32 plMP</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetMarkupPosition(out Int32 plMP)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			plMP = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(plMP);
			object returnItem = Invoker.MethodReturn(this, "GetMarkupPosition", paramsArray, modifiers);
			plMP = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pContainer">NetOffice.MSHTMLApi.IMarkupContainer pContainer</param>
		/// <param name="lMP">Int32 lMP</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToMarkupPosition(NetOffice.MSHTMLApi.IMarkupContainer pContainer, Int32 lMP)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToMarkupPosition", pContainer, lMP);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="muAction">NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction</param>
		/// <param name="pIBoundary">NetOffice.MSHTMLApi.IMarkupPointer pIBoundary</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveUnitBounded(NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction, NetOffice.MSHTMLApi.IMarkupPointer pIBoundary)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUnitBounded", muAction, pIBoundary);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pRight">NetOffice.MSHTMLApi.IMarkupPointer pRight</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsInsideURL(NetOffice.MSHTMLApi.IMarkupPointer pRight, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pRight, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsInsideURL", paramsArray, modifiers);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="fAtStart">Int32 fAtStart</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToContent(NetOffice.MSHTMLApi.IHTMLElement pIElement, Int32 fAtStart)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToContent", pIElement, fAtStart);
		}

		#endregion

		#pragma warning restore
	}
}

