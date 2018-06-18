using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupPointer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F49F-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IMarkupPointer : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppDoc">NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 OwningDoc(out NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY pGravity</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Gravity(out NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY pGravity);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="gravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY gravity</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetGravity(NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY gravity);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfCling">Int32 pfCling</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Cling(out Int32 pfCling);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fCLing">Int32 fCLing</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCling(Int32 fCLing);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 Unposition();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfPositioned">Int32 pfPositioned</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsPositioned(out Int32 pfPositioned);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppContainer">NetOffice.MSHTMLApi.IMarkupContainer ppContainer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppContainer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElement">NetOffice.MSHTMLApi.IHTMLElement pElement</param>
		/// <param name="eAdj">NetOffice.MSHTMLApi.Enums._ELEMENT_ADJACENCY eAdj</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveAdjacentToElement(NetOffice.MSHTMLApi.IHTMLElement pElement, NetOffice.MSHTMLApi.Enums._ELEMENT_ADJACENCY eAdj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pContainer">NetOffice.MSHTMLApi.IMarkupContainer pContainer</param>
		/// <param name="fAtStart">Int32 fAtStart</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToContainer(NetOffice.MSHTMLApi.IMarkupContainer pContainer, Int32 fAtStart);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fMove">Int32 fMove</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		/// <param name="pcch">Int32 pcch</param>
		/// <param name="pchText">Int16 pchText</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 left(Int32 fMove, out NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext, out NetOffice.MSHTMLApi.IHTMLElement ppElement, Int32 pcch, out Int16 pchText);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fMove">Int32 fMove</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext</param>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		/// <param name="pcch">Int32 pcch</param>
		/// <param name="pchText">Int16 pchText</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 right(Int32 fMove, out NetOffice.MSHTMLApi.Enums._MARKUP_CONTEXT_TYPE pContext, out NetOffice.MSHTMLApi.IHTMLElement ppElement, Int32 pcch, out Int16 pchText);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElemCurrent">NetOffice.MSHTMLApi.IHTMLElement ppElemCurrent</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CurrentScope(out NetOffice.MSHTMLApi.IHTMLElement ppElemCurrent);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsLeftOf(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsLeftOfOrEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsRightOf(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsRightOfOrEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerThat">NetOffice.MSHTMLApi.IMarkupPointer pPointerThat</param>
		/// <param name="pfAreEqual">Int32 pfAreEqual</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsEqualTo(NetOffice.MSHTMLApi.IMarkupPointer pPointerThat, out Int32 pfAreEqual);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="muAction">NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveUnit(NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchFindText">string pchFindText</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="pIEndMatch">NetOffice.MSHTMLApi.IMarkupPointer pIEndMatch</param>
		/// <param name="pIEndSearch">NetOffice.MSHTMLApi.IMarkupPointer pIEndSearch</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 findText(string pchFindText, Int32 dwFlags, NetOffice.MSHTMLApi.IMarkupPointer pIEndMatch, NetOffice.MSHTMLApi.IMarkupPointer pIEndSearch);

		#endregion
	}
}
