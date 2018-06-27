using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupPointer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IMarkupPointer : COMObject, NetOffice.MSHTMLApi.IMarkupPointer
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupPointer);
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
                    _type = typeof(IMarkupPointer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupPointer() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppDoc">NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 OwningDoc(out NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppDoc = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppDoc);
			object returnItem = Invoker.MethodReturn(this, "OwningDoc", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLDocument2>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLDocument2));
            else
                ppDoc = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY pGravity</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Gravity(out NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY pGravity)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pGravity = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pGravity);
			object returnItem = Invoker.MethodReturn(this, "Gravity", paramsArray), modifier;
			pGravity = (NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="gravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY gravity</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetGravity(NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY gravity)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetGravity", gravity);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfCling">Int32 pfCling</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Cling(out Int32 pfCling)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfCling = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfCling);
			object returnItem = Invoker.MethodReturn(this, "Cling", paramsArray, modifiers);
			pfCling = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fCLing">Int32 fCLing</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCling(Int32 fCLing)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCling", fCLing);
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
		/// <param name="ppContainer">NetOffice.MSHTMLApi.IMarkupContainer ppContainer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppContainer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppContainer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppContainer);
			object returnItem = Invoker.MethodReturn(this, "GetContainer", paramsArray, modifiers);
			ppContainer = (NetOffice.MSHTMLApi.IMarkupContainer)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="eAdj">NetOffice.MSHTMLApi.Enums._ELEMENT_ADJACENCY eAdj</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveAdjacentToElement(NetOffice.MSHTMLApi.IHTMLElement pElement, NetOffice.MSHTMLApi.Enums._ELEMENT_ADJACENCY eAdj)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveAdjacentToElement", pElement, eAdj);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointer)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToPointer", pPointer);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pContainer">NetOffice.MSHTMLApi.IMarkupContainer pContainer</param>
		/// <param name="fAtStart">Int32 fAtStart</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToContainer(NetOffice.MSHTMLApi.IMarkupContainer pContainer, Int32 fAtStart)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToContainer", pContainer, fAtStart);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fMove">Int32 fMove</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		/// <param name="pcch">Int32 pcch</param>
		/// <param name="pchText">Int16 pchText</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 left(Int32 fMove, out NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext, out NetOffice.MSHTMLApi.IHTMLElement ppElement, Int32 pcch, out Int16 pchText)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,false,true);
			pContext = 0;
			ppElement = null;
			pchText = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(fMove, pContext, ppElement, pcch, pchText);
			object returnItem = Invoker.MethodReturn(this, "left", paramsArray, modifiers);
			pContext = (NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE)paramsArray[1];
            if (paramsArray[2] is MarshalByRefObject)
                ppElement = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[2], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElement = null;
            pchText = (Int16)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fMove">Int32 fMove</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		/// <param name="pcch">Int32 pcch</param>
		/// <param name="pchText">Int16 pchText</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 right(Int32 fMove, out NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext, out NetOffice.MSHTMLApi.IHTMLElement ppElement, Int32 pcch, out Int16 pchText)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,false,true);
			pContext = 0;
			ppElement = null;
			pchText = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(fMove, pContext, ppElement, pcch, pchText);
			object returnItem = Invoker.MethodReturn(this, "right", paramsArray, modifiers);
			pContext = (NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE)paramsArray[1];
            if (paramsArray[2] is MarshalByRefObject)
                ppElement = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[2], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElement = null;
			pchText = (Int16)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElemCurrent">NetOffice.MSHTMLApi.IHTMLElement ppElemCurrent</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CurrentScope(out NetOffice.MSHTMLApi.IHTMLElement ppElemCurrent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppElemCurrent = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppElemCurrent);
			object returnItem = Invoker.MethodReturn(this, "CurrentScope", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppElemCurrent = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElemCurrent = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsLeftOf(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsLeftOf", paramsArray, modifiers);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsLeftOfOrEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsLeftOfOrEqualTo", paramsArray, modifiers);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsRightOf(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsRightOf", paramsArray, modifiers);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsRightOfOrEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsRightOfOrEqualTo", paramsArray, modifiers);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfAreEqual">Int32 pfAreEqual</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfAreEqual)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfAreEqual = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfAreEqual);
			object returnItem = Invoker.MethodReturn(this, "IsEqualTo", paramsArray, modifiers);
			pfAreEqual = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="muAction">NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveUnit(NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUnit", muAction);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchFindText">string pchFindText</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="pIEndMatch">NetOffice.MSHTMLApi.IMarkupPointer pIEndMatch</param>
		/// <param name="pIEndSearch">NetOffice.MSHTMLApi.IMarkupPointer pIEndSearch</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 findText(string pchFindText, Int32 dwFlags, NetOffice.MSHTMLApi.IMarkupPointer pIEndMatch, NetOffice.MSHTMLApi.IMarkupPointer pIEndSearch)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "findText", pchFindText, dwFlags, pIEndMatch, pIEndSearch);
		}

		#endregion

		#pragma warning restore
	}
}

