using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLPainter 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6A6-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLPainter : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="rcBounds">tagRECT rcBounds</param>
		/// <param name="rcUpdate">tagRECT rcUpdate</param>
		/// <param name="lDrawFlags">Int32 lDrawFlags</param>
		/// <param name="hdc">_RemotableHandle hdc</param>
		/// <param name="pvDrawObject">object pvDrawObject</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Draw(tagRECT rcBounds, tagRECT rcUpdate, Int32 lDrawFlags, _RemotableHandle hdc, object pvDrawObject);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="size">tagSIZE size</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 onresize(tagSIZE size);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pInfo">_HTML_PAINTER_INFO pInfo</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetPainterInfo(out _HTML_PAINTER_INFO pInfo);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pt">tagPOINT pt</param>
		/// <param name="pbHit">Int32 pbHit</param>
		/// <param name="plPartID">Int32 plPartID</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 HitTestPoint(tagPOINT pt, out Int32 pbHit, out Int32 plPartID);

		#endregion
	}
}
