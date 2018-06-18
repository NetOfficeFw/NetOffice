using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLCaret 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F604-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLCaret : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveCaretToPointer(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, Int32 fScrollIntoView, NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="fVisible">Int32 fVisible</param>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveCaretToPointerEx(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, Int32 fVisible, Int32 fScrollIntoView, NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIMarkupPointer">NetOffice.MSHTMLApi.IMarkupPointer pIMarkupPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveMarkupPointerToCaret(NetOffice.MSHTMLApi.IMarkupPointer pIMarkupPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveDisplayPointerToCaret(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIsVisible">Int32 pIsVisible</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsVisible(out Int32 pIsVisible);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Show(Int32 fScrollIntoView);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 Hide();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pText">Int16 pText</param>
		/// <param name="lLen">Int32 lLen</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InsertText(Int16 pText, Int32 lLen);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 scrollIntoView();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="fTranslate">Int32 fTranslate</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetLocation(out tagPOINT pPoint, Int32 fTranslate);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION peDir</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCaretDirection(out NetOffice.MSHTMLApi.Enums._CARET_DIRECTION peDir);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCaretDirection(NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir);

		#endregion
	}
}
