using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IDisplayPointer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F69E-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IDisplayPointer : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ptPoint">tagPOINT ptPoint</param>
		/// <param name="eCoordSystem">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eCoordSystem</param>
		/// <param name="pElementContext">NetOffice.MSHTMLApi.IHTMLElement pElementContext</param>
		/// <param name="dwHitTestOptions">Int32 dwHitTestOptions</param>
		/// <param name="pdwHitTestResults">Int32 pdwHitTestResults</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 moveToPoint(tagPOINT ptPoint, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eCoordSystem, NetOffice.MSHTMLApi.IHTMLElement pElementContext, Int32 dwHitTestOptions, out Int32 pdwHitTestResults);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eMoveUnit">NetOffice.MSHTMLApi.Enums._DISPLAY_MOVEUNIT eMoveUnit</param>
		/// <param name="lXPos">Int32 lXPos</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveUnit(NetOffice.MSHTMLApi.Enums._DISPLAY_MOVEUNIT eMoveUnit, Int32 lXPos);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pMarkupPointer">NetOffice.MSHTMLApi.IMarkupPointer pMarkupPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 PositionMarkupPointer(NetOffice.MSHTMLApi.IMarkupPointer pMarkupPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToPointer(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY eGravity</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetPointerGravity(NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY eGravity);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peGravity">NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY peGravity</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetPointerGravity(out NetOffice.MSHTMLApi.Enums._POINTER_GRAVITY peGravity);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eGravity">NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY eGravity</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetDisplayGravity(NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY eGravity);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peGravity">NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY peGravity</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetDisplayGravity(out NetOffice.MSHTMLApi.Enums._DISPLAY_GRAVITY peGravity);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfPositioned">Int32 pfPositioned</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsPositioned(out Int32 pfPositioned);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 Unposition();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="pfIsEqual">Int32 pfIsEqual</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsEqualTo(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, out Int32 pfIsEqual);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="pfIsLeftOf">Int32 pfIsLeftOf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsLeftOf(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, out Int32 pfIsLeftOf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="pfIsRightOf">Int32 pfIsRightOf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsRightOf(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, out Int32 pfIsRightOf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfBOL">Int32 pfBOL</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsAtBOL(out Int32 pfBOL);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		/// <param name="pDispLineContext">NetOffice.MSHTMLApi.IDisplayPointer pDispLineContext</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToMarkupPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointer, NetOffice.MSHTMLApi.IDisplayPointer pDispLineContext);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 scrollIntoView();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppLineInfo">NetOffice.MSHTMLApi.ILineInfo ppLineInfo</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetLineInfo(out NetOffice.MSHTMLApi.ILineInfo ppLineInfo);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppLayoutElement">NetOffice.MSHTMLApi.IHTMLElement ppLayoutElement</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetFlowElement(out NetOffice.MSHTMLApi.IHTMLElement ppLayoutElement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pdwBreaks">Int32 pdwBreaks</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 QueryBreaks(out Int32 pdwBreaks);

		#endregion
	}
}
