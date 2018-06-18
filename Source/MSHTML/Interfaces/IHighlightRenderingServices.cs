using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHighlightRenderingServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F606-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHighlightRenderingServices : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointerStart">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart</param>
		/// <param name="pDispPointerEnd">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd</param>
		/// <param name="pIRenderStyle">NetOffice.MSHTMLApi.IHTMLRenderStyle pIRenderStyle</param>
		/// <param name="ppISegment">NetOffice.MSHTMLApi.IHighlightSegment ppISegment</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AddSegment(NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd, NetOffice.MSHTMLApi.IHTMLRenderStyle pIRenderStyle, out NetOffice.MSHTMLApi.IHighlightSegment ppISegment);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.IHighlightSegment pISegment</param>
		/// <param name="pDispPointerStart">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart</param>
		/// <param name="pDispPointerEnd">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveSegmentToPointers(NetOffice.MSHTMLApi.IHighlightSegment pISegment, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.IHighlightSegment pISegment</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RemoveSegment(NetOffice.MSHTMLApi.IHighlightSegment pISegment);

		#endregion
	}
}
