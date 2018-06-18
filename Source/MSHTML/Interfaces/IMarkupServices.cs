using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F4A0-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IMarkupServices : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppPointer">NetOffice.MSHTMLApi.IMarkupPointer ppPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateMarkupPointer(out NetOffice.MSHTMLApi.IMarkupPointer ppPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppMarkupContainer">NetOffice.MSHTMLApi.IMarkupContainer ppMarkupContainer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateMarkupContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppMarkupContainer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagID">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID</param>
		/// <param name="pchAttributes">Int16 pchAttributes</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 createElement(NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID, Int16 pchAttributes, out NetOffice.MSHTMLApi.IHTMLElement ppElement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElemCloneThis">NetOffice.MSHTMLApi.IHTMLElement pElemCloneThis</param>
		/// <param name="ppElementTheClone">NetOffice.MSHTMLApi.IHTMLElement ppElementTheClone</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CloneElement(NetOffice.MSHTMLApi.IHTMLElement pElemCloneThis, out NetOffice.MSHTMLApi.IHTMLElement ppElementTheClone);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElementInsert">NetOffice.MSHTMLApi.IHTMLElement pElementInsert</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InsertElement(NetOffice.MSHTMLApi.IHTMLElement pElementInsert, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElementRemove">NetOffice.MSHTMLApi.IHTMLElement pElementRemove</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RemoveElement(NetOffice.MSHTMLApi.IHTMLElement pElementRemove);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 remove(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerSourceStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart</param>
		/// <param name="pPointerSourceFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Copy(NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerSourceStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart</param>
		/// <param name="pPointerSourceFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 move(NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerSourceFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchText">Int16 pchText</param>
		/// <param name="cch">Int32 cch</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InsertText(Int16 pchText, Int32 cch, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchHTML">Int16 pchHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="ppPointerStart">NetOffice.MSHTMLApi.IMarkupPointer ppPointerStart</param>
		/// <param name="ppPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer ppPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ParseString(Int16 pchHTML, Int32 dwFlags, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer ppPointerStart, NetOffice.MSHTMLApi.IMarkupPointer ppPointerFinish);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hglobalHTML">_userHGLOBAL hglobalHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ParseGlobal(_userHGLOBAL hglobalHTML, Int32 dwFlags, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="pfScoped">Int32 pfScoped</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsScopedElement(NetOffice.MSHTMLApi.IHTMLElement pElement, out Int32 pfScoped);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="ptagId">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetElementTagId(NetOffice.MSHTMLApi.IHTMLElement pElement, out NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="ptagId">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetTagIDForName(string bstrName, out NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID ptagId);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagID">NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetNameForTagID(NetOffice.MSHTMLApi.Enums._ELEMENT_TAG_ID tagID, out string pbstrName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIRange">NetOffice.MSHTMLApi.IHTMLTxtRange pIRange</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MovePointersToRange(NetOffice.MSHTMLApi.IHTMLTxtRange pIRange, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		/// <param name="pIRange">NetOffice.MSHTMLApi.IHTMLTxtRange pIRange</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveRangeToPointers(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish, NetOffice.MSHTMLApi.IHTMLTxtRange pIRange);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchTitle">Int16 pchTitle</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 BeginUndoUnit(Int16 pchTitle);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 EndUndoUnit();

		#endregion
	}
}
