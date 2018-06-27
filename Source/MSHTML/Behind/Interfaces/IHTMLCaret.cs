using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLCaret 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLCaret : COMObject, NetOffice.MSHTMLApi.IHTMLCaret
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLCaret);
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
                    _type = typeof(IHTMLCaret);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLCaret() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveCaretToPointer(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, Int32 fScrollIntoView, NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveCaretToPointer", pDispPointer, fScrollIntoView, eDir);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="fVisible">Int32 fVisible</param>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveCaretToPointerEx(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, Int32 fVisible, Int32 fScrollIntoView, NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveCaretToPointerEx", pDispPointer, fVisible, fScrollIntoView, eDir);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIMarkupPointer">NetOffice.MSHTMLApi.IMarkupPointer pIMarkupPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveMarkupPointerToCaret(NetOffice.MSHTMLApi.IMarkupPointer pIMarkupPointer)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveMarkupPointerToCaret", pIMarkupPointer);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveDisplayPointerToCaret(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveDisplayPointerToCaret", pDispPointer);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIsVisible">Int32 pIsVisible</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsVisible(out Int32 pIsVisible)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pIsVisible = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pIsVisible);
			object returnItem = Invoker.MethodReturn(this, "IsVisible", paramsArray, modifiers);
			pIsVisible = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Show(Int32 fScrollIntoView)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Show", fScrollIntoView);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Hide()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Hide");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pText">Int16 pText</param>
		/// <param name="lLen">Int32 lLen</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InsertText(Int16 pText, Int32 lLen)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InsertText", pText, lLen);
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
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="fTranslate">Int32 fTranslate</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetLocation(out tagPOINT pPoint, Int32 fTranslate)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			pPoint = new NetOffice.MSHTMLApi.tagPOINT();
			object[] paramsArray = Invoker.ValidateParamsArray(pPoint, fTranslate);
			object returnItem = Invoker.MethodReturn(this, "GetLocation", paramsArray, modifiers);
			pPoint = (tagPOINT)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION peDir</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCaretDirection(out NetOffice.MSHTMLApi.Enums._CARET_DIRECTION peDir)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			peDir = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(peDir);
			object returnItem = Invoker.MethodReturn(this, "GetCaretDirection", paramsArray, modifiers);
			peDir = (NetOffice.MSHTMLApi.Enums._CARET_DIRECTION)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCaretDirection(NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCaretDirection", eDir);
		}

		#endregion

		#pragma warning restore
	}
}

