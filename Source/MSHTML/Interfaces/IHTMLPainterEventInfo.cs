using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLPainterEventInfo 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6DF-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLPainterEventInfo : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plEventInfoFlags">Int32 plEventInfoFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetEventInfoFlags(out Int32 plEventInfoFlags);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetEventTarget(NetOffice.MSHTMLApi.IHTMLElement ppElement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lPartID">Int32 lPartID</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCursor(Int32 lPartID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lPartID">Int32 lPartID</param>
		/// <param name="pbstrPart">string pbstrPart</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 StringFromPartID(Int32 lPartID, out string pbstrPart);

		#endregion
	}
}
