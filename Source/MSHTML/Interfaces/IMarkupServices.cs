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
	/// Interface IMarkupServices 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IMarkupServices : COMObject
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
                    _type = typeof(IMarkupServices);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMarkupServices(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupServices(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupServices(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupServices(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupServices(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupServices() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupServices(string progId) : base(progId)
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
		/// <param name="ppPointer">NetOffice.MSHTMLApi.IMarkupPointer ppPointer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CreateMarkupPointer(out NetOffice.MSHTMLApi.IMarkupPointer ppPointer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppPointer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppPointer);
			object returnItem = Invoker.MethodReturn(this, "CreateMarkupPointer", paramsArray);
			ppPointer = (NetOffice.MSHTMLApi.IMarkupPointer)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppMarkupContainer">NetOffice.MSHTMLApi.IMarkupContainer ppMarkupContainer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CreateMarkupContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppMarkupContainer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppMarkupContainer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppMarkupContainer);
			object returnItem = Invoker.MethodReturn(this, "CreateMarkupContainer", paramsArray);
			ppMarkupContainer = (NetOffice.MSHTMLApi.IMarkupContainer)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="tagID">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID</param>
		/// <param name="pchAttributes">Int16 pchAttributes</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 createElement(NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID, Int16 pchAttributes, out NetOffice.MSHTMLApi.IHTMLElement ppElement)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			ppElement = null;
			object[] paramsArray = Invoker.ValidateParamsArray(tagID, pchAttributes, ppElement);
			object returnItem = Invoker.MethodReturn(this, "createElement", paramsArray);
			ppElement = (NetOffice.MSHTMLApi.IHTMLElement)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pElemCloneThis">NetOffice.MSHTMLApi.IHTMLElement pElemCloneThis</param>
		/// <param name="ppElementTheClone">NetOffice.MSHTMLApi.IHTMLElement ppElementTheClone</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CloneElement(NetOffice.MSHTMLApi.IHTMLElement pElemCloneThis, out NetOffice.MSHTMLApi.IHTMLElement ppElementTheClone)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppElementTheClone = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pElemCloneThis, ppElementTheClone);
			object returnItem = Invoker.MethodReturn(this, "CloneElement", paramsArray);
			ppElementTheClone = (NetOffice.MSHTMLApi.IHTMLElement)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pElementInsert">NetOffice.MSHTMLApi.IHTMLElement pElementInsert</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 InsertElement(NetOffice.MSHTMLApi.IHTMLElement pElementInsert, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pElementInsert, pPointerStart, pPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "InsertElement", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pElementRemove">NetOffice.MSHTMLApi.IHTMLElement pElementRemove</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RemoveElement(NetOffice.MSHTMLApi.IHTMLElement pElementRemove)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pElementRemove);
			object returnItem = Invoker.MethodReturn(this, "RemoveElement", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 remove(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerStart, pPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "remove", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerSourceStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart</param>
		/// <param name="pPointerSourceFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Copy(NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerSourceStart, pPointerSourceFinish, pPointerTarget);
			object returnItem = Invoker.MethodReturn(this, "Copy", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerSourceStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart</param>
		/// <param name="pPointerSourceFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 move(NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerSourceStart, pPointerSourceFinish, pPointerTarget);
			object returnItem = Invoker.MethodReturn(this, "move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pchText">Int16 pchText</param>
		/// <param name="cch">Int32 cch</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 InsertText(Int16 pchText, Int32 cch, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pchText, cch, pPointerTarget);
			object returnItem = Invoker.MethodReturn(this, "InsertText", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pchHTML">Int16 pchHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="ppPointerStart">NetOffice.MSHTMLApi.IMarkupPointer ppPointerStart</param>
		/// <param name="ppPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer ppPointerFinish</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 ParseString(Int16 pchHTML, Int32 dwFlags, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer ppPointerStart, NetOffice.MSHTMLApi.IMarkupPointer ppPointerFinish)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,false,false);
			ppContainerResult = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pchHTML, dwFlags, ppContainerResult, ppPointerStart, ppPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "ParseString", paramsArray);
			ppContainerResult = (NetOffice.MSHTMLApi.IMarkupContainer)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hglobalHTML">_userHGLOBAL hglobalHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 ParseGlobal(_userHGLOBAL hglobalHTML, Int32 dwFlags, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,false,false);
			ppContainerResult = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hglobalHTML, dwFlags, ppContainerResult, pPointerStart, pPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "ParseGlobal", paramsArray);
			ppContainerResult = (NetOffice.MSHTMLApi.IMarkupContainer)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="pfScoped">Int32 pfScoped</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsScopedElement(NetOffice.MSHTMLApi.IHTMLElement pElement, out Int32 pfScoped)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfScoped = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pElement, pfScoped);
			object returnItem = Invoker.MethodReturn(this, "IsScopedElement", paramsArray);
			pfScoped = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="ptagId">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetElementTagId(NetOffice.MSHTMLApi.IHTMLElement pElement, out NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ptagId = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pElement, ptagId);
			object returnItem = Invoker.MethodReturn(this, "GetElementTagId", paramsArray);
			ptagId = (NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="ptagId">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetTagIDForName(string bstrName, out NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ptagId = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrName, ptagId);
			object returnItem = Invoker.MethodReturn(this, "GetTagIDForName", paramsArray);
			ptagId = (NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="tagID">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetNameForTagID(NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID, out string pbstrName)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(tagID, pbstrName);
			object returnItem = Invoker.MethodReturn(this, "GetNameForTagID", paramsArray);
			pbstrName = (string)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIRange">NetOffice.MSHTMLApi.IHTMLTxtRange pIRange</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MovePointersToRange(NetOffice.MSHTMLApi.IHTMLTxtRange pIRange, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIRange, pPointerStart, pPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "MovePointersToRange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		/// <param name="pIRange">NetOffice.MSHTMLApi.IHTMLTxtRange pIRange</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveRangeToPointers(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish, NetOffice.MSHTMLApi.IHTMLTxtRange pIRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerStart, pPointerFinish, pIRange);
			object returnItem = Invoker.MethodReturn(this, "MoveRangeToPointers", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pchTitle">Int16 pchTitle</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 BeginUndoUnit(Int16 pchTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pchTitle);
			object returnItem = Invoker.MethodReturn(this, "BeginUndoUnit", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EndUndoUnit()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "EndUndoUnit", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}