using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorRender 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F4AA-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorRender : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hdc">_RemotableHandle hdc</param>
		/// <param name="lLayer">Int32 lLayer</param>
		/// <param name="pRect">tagRECT pRect</param>
		/// <param name="pReserved">object pReserved</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Draw(_RemotableHandle hdc, Int32 lLayer, tagRECT pRect, object pReserved);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetRenderInfo();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="pReserved">object pReserved</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 HitTestPoint(tagPOINT pPoint, object pReserved);

		#endregion
	}
}
