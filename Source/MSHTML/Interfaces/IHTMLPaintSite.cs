using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLPaintSite 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6A7-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLPaintSite : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 InvalidatePainterInfo();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="prcInvalid">tagRECT prcInvalid</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InvalidateRect(tagRECT prcInvalid);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="rgnInvalid">_RemotableHandle rgnInvalid</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InvalidateRegion(_RemotableHandle rgnInvalid);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pDrawInfo">_HTML_PAINT_DRAW_INFO pDrawInfo</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetDrawInfo(Int32 lFlags, out _HTML_PAINT_DRAW_INFO pDrawInfo);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ptGlobal">tagPOINT ptGlobal</param>
		/// <param name="pptLocal">tagPOINT pptLocal</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 TransformGlobalToLocal(tagPOINT ptGlobal, out tagPOINT pptLocal);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ptLocal">tagPOINT ptLocal</param>
		/// <param name="pptGlobal">tagPOINT pptGlobal</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 TransformLocalToGlobal(tagPOINT ptLocal, out tagPOINT pptGlobal);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plCookie">Int32 plCookie</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetHitTestCookie(out Int32 plCookie);

		#endregion
	}
}
