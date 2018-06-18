using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupPointer2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F675-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IMarkupPointer2 : IMarkupPointer
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfAtBreak">Int32 pfAtBreak</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsAtWordBreak(out Int32 pfAtBreak);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plMP">Int32 plMP</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetMarkupPosition(out Int32 plMP);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pContainer">NetOffice.MSHTMLApi.IMarkupContainer pContainer</param>
		/// <param name="lMP">Int32 lMP</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToMarkupPosition(NetOffice.MSHTMLApi.IMarkupContainer pContainer, Int32 lMP);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="muAction">NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction</param>
		/// <param name="pIBoundary">NetOffice.MSHTMLApi.IMarkupPointer pIBoundary</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveUnitBounded(NetOffice.MSHTMLApi.Enums._MOVEUNIT_ACTION muAction, NetOffice.MSHTMLApi.IMarkupPointer pIBoundary);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pRight">NetOffice.MSHTMLApi.IMarkupPointer pRight</param>
		/// <param name="pfResult">Int32 pfResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsInsideURL(NetOffice.MSHTMLApi.IMarkupPointer pRight, out Int32 pfResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="fAtStart">Int32 fAtStart</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToContent(NetOffice.MSHTMLApi.IHTMLElement pIElement, Int32 fAtStart);

		#endregion
	}
}
