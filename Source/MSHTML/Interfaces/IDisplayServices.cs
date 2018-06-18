using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IDisplayServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F69D-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IDisplayServices : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppDispPointer">NetOffice.MSHTMLApi.IDisplayPointer ppDispPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateDisplayPointer(out NetOffice.MSHTMLApi.IDisplayPointer ppDispPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pRect">tagRECT pRect</param>
		/// <param name="eSource">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource</param>
		/// <param name="eDestination">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination</param>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 TransformRect(tagRECT pRect, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination, NetOffice.MSHTMLApi.IHTMLElement pIElement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="eSource">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource</param>
		/// <param name="eDestination">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination</param>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 TransformPoint(tagPOINT pPoint, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination, NetOffice.MSHTMLApi.IHTMLElement pIElement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppCaret">NetOffice.MSHTMLApi.IHTMLCaret ppCaret</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCaret(out NetOffice.MSHTMLApi.IHTMLCaret ppCaret);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		/// <param name="ppComputedStyle">NetOffice.MSHTMLApi.IHTMLComputedStyle ppComputedStyle</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetComputedStyle(NetOffice.MSHTMLApi.IMarkupPointer pPointer, out NetOffice.MSHTMLApi.IHTMLComputedStyle ppComputedStyle);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="rect">tagRECT rect</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ScrollRectIntoView(NetOffice.MSHTMLApi.IHTMLElement pIElement, tagRECT rect);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="pfHasFlowLayout">Int32 pfHasFlowLayout</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 HasFlowLayout(NetOffice.MSHTMLApi.IHTMLElement pIElement, out Int32 pfHasFlowLayout);

		#endregion
	}
}
