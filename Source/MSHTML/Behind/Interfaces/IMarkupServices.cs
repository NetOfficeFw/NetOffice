using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IMarkupServices : COMObject, NetOffice.MSHTMLApi.IMarkupServices
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupServices);
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
                    _type = typeof(IMarkupServices);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupServices() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppPointer">NetOffice.MSHTMLApi.IMarkupPointer ppPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CreateMarkupPointer(out NetOffice.MSHTMLApi.IMarkupPointer ppPointer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppPointer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppPointer);
			object returnItem = Invoker.MethodReturn(this, "CreateMarkupPointer", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppPointer = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IMarkupPointer>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IMarkupPointer));
            else
                ppPointer = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppMarkupContainer">NetOffice.MSHTMLApi.IMarkupContainer ppMarkupContainer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CreateMarkupContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppMarkupContainer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppMarkupContainer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppMarkupContainer);
			object returnItem = Invoker.MethodReturn(this, "CreateMarkupContainer", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppMarkupContainer = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IMarkupContainer>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IMarkupContainer));
            else
                ppMarkupContainer = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagID">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID</param>
		/// <param name="pchAttributes">Int16 pchAttributes</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 createElement(NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID, Int16 pchAttributes, out NetOffice.MSHTMLApi.IHTMLElement ppElement)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			ppElement = null;
			object[] paramsArray = Invoker.ValidateParamsArray(tagID, pchAttributes, ppElement);
			object returnItem = Invoker.MethodReturn(this, "createElement", paramsArray, modifiers);
            if (paramsArray[2] is MarshalByRefObject)
                ppElement = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[2], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElement = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElemCloneThis">NetOffice.MSHTMLApi.IHTMLElement pElemCloneThis</param>
		/// <param name="ppElementTheClone">NetOffice.MSHTMLApi.IHTMLElement ppElementTheClone</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CloneElement(NetOffice.MSHTMLApi.IHTMLElement pElemCloneThis, out NetOffice.MSHTMLApi.IHTMLElement ppElementTheClone)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppElementTheClone = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pElemCloneThis, ppElementTheClone);
			object returnItem = Invoker.MethodReturn(this, "CloneElement", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppElementTheClone = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[1], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElementTheClone = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElementInsert">NetOffice.MSHTMLApi.IHTMLElement pElementInsert</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InsertElement(NetOffice.MSHTMLApi.IHTMLElement pElementInsert, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InsertElement", pElementInsert, pPointerStart, pPointerFinish);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElementRemove">NetOffice.MSHTMLApi.IHTMLElement pElementRemove</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RemoveElement(NetOffice.MSHTMLApi.IHTMLElement pElementRemove)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveElement", pElementRemove);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 remove(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "remove", pPointerStart, pPointerFinish);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerSourceStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart</param>
		/// <param name="pPointerSourceFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Copy(NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Copy", pPointerSourceStart, pPointerSourceFinish, pPointerTarget);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerSourceStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart</param>
		/// <param name="pPointerSourceFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 move(NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "move", pPointerSourceStart, pPointerSourceFinish, pPointerTarget);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchText">Int16 pchText</param>
		/// <param name="cch">Int32 cch</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InsertText(Int16 pchText, Int32 cch, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InsertText", pchText, cch, pPointerTarget);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchHTML">Int16 pchHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="ppPointerStart">NetOffice.MSHTMLApi.IMarkupPointer ppPointerStart</param>
		/// <param name="ppPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer ppPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ParseString(Int16 pchHTML, Int32 dwFlags, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer ppPointerStart, NetOffice.MSHTMLApi.IMarkupPointer ppPointerFinish)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,false,false);
			ppContainerResult = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pchHTML, dwFlags, ppContainerResult, ppPointerStart, ppPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "ParseString", paramsArray, modifiers);
            if (paramsArray[2] is MarshalByRefObject)
                ppContainerResult = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IMarkupContainer>(this, paramsArray[2], typeof(NetOffice.MSHTMLApi.IMarkupContainer));
            else
                ppContainerResult = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hglobalHTML">_userHGLOBAL hglobalHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ParseGlobal(_userHGLOBAL hglobalHTML, Int32 dwFlags, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,false,false);
			ppContainerResult = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hglobalHTML, dwFlags, ppContainerResult, pPointerStart, pPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "ParseGlobal", paramsArray, modifiers);
            if (paramsArray[2] is MarshalByRefObject)
                ppContainerResult = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IMarkupContainer>(this, paramsArray[2], typeof(NetOffice.MSHTMLApi.IMarkupContainer));
            else
                ppContainerResult = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="pfScoped">Int32 pfScoped</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsScopedElement(NetOffice.MSHTMLApi.IHTMLElement pElement, out Int32 pfScoped)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfScoped = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pElement, pfScoped);
			object returnItem = Invoker.MethodReturn(this, "IsScopedElement", paramsArray, modifiers);
			pfScoped = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="ptagId">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetElementTagId(NetOffice.MSHTMLApi.IHTMLElement pElement, out NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ptagId = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pElement, ptagId);
			object returnItem = Invoker.MethodReturn(this, "GetElementTagId", paramsArray, modifiers);
			ptagId = (NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="ptagId">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetTagIDForName(string bstrName, out NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ptagId = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrName, ptagId);
			object returnItem = Invoker.MethodReturn(this, "GetTagIDForName", paramsArray, modifiers);
			ptagId = (NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagID">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetNameForTagID(NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID, out string pbstrName)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(tagID, pbstrName);
			object returnItem = Invoker.MethodReturn(this, "GetNameForTagID", paramsArray, modifiers);
			pbstrName = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIRange">NetOffice.MSHTMLApi.IHTMLTxtRange pIRange</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MovePointersToRange(NetOffice.MSHTMLApi.IHTMLTxtRange pIRange, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MovePointersToRange", pIRange, pPointerStart, pPointerFinish);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		/// <param name="pIRange">NetOffice.MSHTMLApi.IHTMLTxtRange pIRange</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveRangeToPointers(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish, NetOffice.MSHTMLApi.IHTMLTxtRange pIRange)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveRangeToPointers", pPointerStart, pPointerFinish, pIRange);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchTitle">Int16 pchTitle</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 BeginUndoUnit(Int16 pchTitle)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeginUndoUnit", pchTitle);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 EndUndoUnit()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndUndoUnit");
		}

		#endregion

		#pragma warning restore
	}
}

