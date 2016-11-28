using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// Interface IMarkupPointer 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IMarkupPointer : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMarkupPointer(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupPointer(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupPointer(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupPointer(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupPointer(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupPointer() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupPointer(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppDoc">NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 OwningDoc(out NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppDoc = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppDoc);
			object returnItem = Invoker.MethodReturn(this, "OwningDoc", paramsArray);
			ppDoc = (NetOffice.MSHTMLApi.IHTMLDocument2)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY pGravity</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Gravity(out NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY pGravity)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pGravity = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pGravity);
			object returnItem = Invoker.MethodReturn(this, "Gravity", paramsArray);
			pGravity = (NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="gravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY Gravity</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetGravity(NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY gravity)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gravity);
			object returnItem = Invoker.MethodReturn(this, "SetGravity", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pfCling">Int32 pfCling</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Cling(out Int32 pfCling)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfCling = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfCling);
			object returnItem = Invoker.MethodReturn(this, "Cling", paramsArray);
			pfCling = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fCLing">Int32 fCLing</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCling(Int32 fCLing)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fCLing);
			object returnItem = Invoker.MethodReturn(this, "SetCling", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Unposition()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Unposition", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pfPositioned">Int32 pfPositioned</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsPositioned(out Int32 pfPositioned)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfPositioned = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfPositioned);
			object returnItem = Invoker.MethodReturn(this, "IsPositioned", paramsArray);
			pfPositioned = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppContainer">NetOffice.MSHTMLApi.IMarkupContainer ppContainer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppContainer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppContainer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppContainer);
			object returnItem = Invoker.MethodReturn(this, "GetContainer", paramsArray);
			ppContainer = (NetOffice.MSHTMLApi.IMarkupContainer)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="eAdj">NetOffice.MSHTMLApi.Enums._ELEMENT_ADJACENCY eAdj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveAdjacentToElement(NetOffice.MSHTMLApi.IHTMLElement pElement, NetOffice.MSHTMLApi.Enums._ELEMENT_ADJACENCY eAdj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pElement, eAdj);
			object returnItem = Invoker.MethodReturn(this, "MoveAdjacentToElement", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveToPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pPointer);
			object returnItem = Invoker.MethodReturn(this, "MoveToPointer", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pContainer">NetOffice.MSHTMLApi.IMarkupContainer pContainer</param>
		/// <param name="fAtStart">Int32 fAtStart</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveToContainer(NetOffice.MSHTMLApi.IMarkupContainer pContainer, Int32 fAtStart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pContainer, fAtStart);
			object returnItem = Invoker.MethodReturn(this, "MoveToContainer", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fMove">Int32 fMove</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		/// <param name="pcch">Int32 pcch</param>
		/// <param name="pchText">Int16 pchText</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 left(Int32 fMove, out NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext, out NetOffice.MSHTMLApi.IHTMLElement ppElement, Int32 pcch, out Int16 pchText)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,false,true);
			pContext = 0;
			ppElement = null;
			pchText = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(fMove, pContext, ppElement, pcch, pchText);
			object returnItem = Invoker.MethodReturn(this, "left", paramsArray);
			pContext = (NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE)paramsArray[1];
			ppElement = (NetOffice.MSHTMLApi.IHTMLElement)paramsArray[2];
			pchText = (Int16)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fMove">Int32 fMove</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		/// <param name="pcch">Int32 pcch</param>
		/// <param name="pchText">Int16 pchText</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 right(Int32 fMove, out NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext, out NetOffice.MSHTMLApi.IHTMLElement ppElement, Int32 pcch, out Int16 pchText)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,false,true);
			pContext = 0;
			ppElement = null;
			pchText = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(fMove, pContext, ppElement, pcch, pchText);
			object returnItem = Invoker.MethodReturn(this, "right", paramsArray);
			pContext = (NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE)paramsArray[1];
			ppElement = (NetOffice.MSHTMLApi.IHTMLElement)paramsArray[2];
			pchText = (Int16)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppElemCurrent">NetOffice.MSHTMLApi.IHTMLElement ppElemCurrent</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CurrentScope(out NetOffice.MSHTMLApi.IHTMLElement ppElemCurrent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppElemCurrent = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppElemCurrent);
			object returnItem = Invoker.MethodReturn(this, "CurrentScope", paramsArray);
			ppElemCurrent = (NetOffice.MSHTMLApi.IHTMLElement)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsLeftOf(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsLeftOf", paramsArray);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsLeftOfOrEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsLeftOfOrEqualTo", paramsArray);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsRightOf(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsRightOf", paramsArray);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsRightOfOrEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfResult);
			object returnItem = Invoker.MethodReturn(this, "IsRightOfOrEqualTo", paramsArray);
			pfResult = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfAreEqual">Int32 pfAreEqual</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfAreEqual)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfAreEqual = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerThat, pfAreEqual);
			object returnItem = Invoker.MethodReturn(this, "IsEqualTo", paramsArray);
			pfAreEqual = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="muAction">NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveUnit(NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(muAction);
			object returnItem = Invoker.MethodReturn(this, "MoveUnit", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pchFindText">string pchFindText</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="pIEndMatch">NetOffice.MSHTMLApi.IMarkupPointer pIEndMatch</param>
		/// <param name="pIEndSearch">NetOffice.MSHTMLApi.IMarkupPointer pIEndSearch</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 findText(string pchFindText, Int32 dwFlags, NetOffice.MSHTMLApi.IMarkupPointer pIEndMatch, NetOffice.MSHTMLApi.IMarkupPointer pIEndSearch)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pchFindText, dwFlags, pIEndMatch, pIEndSearch);
			object returnItem = Invoker.MethodReturn(this, "findText", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}