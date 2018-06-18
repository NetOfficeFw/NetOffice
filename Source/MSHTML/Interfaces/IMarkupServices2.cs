using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupServices2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F682-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IMarkupServices2 : IMarkupServices
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hglobalHTML">_userHGLOBAL hglobalHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.IMarkupContainer pContext</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ParseGlobalEx(_userHGLOBAL hglobalHTML, Int32 dwFlags, NetOffice.MSHTMLApi.IMarkupContainer pContext, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		/// <param name="pPointerStatus">NetOffice.MSHTMLApi.IMarkupPointer pPointerStatus</param>
		/// <param name="ppElemFailBottom">NetOffice.MSHTMLApi.IHTMLElement ppElemFailBottom</param>
		/// <param name="ppElemFailTop">NetOffice.MSHTMLApi.IHTMLElement ppElemFailTop</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ValidateElements(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget, NetOffice.MSHTMLApi.IMarkupPointer pPointerStatus, out NetOffice.MSHTMLApi.IHTMLElement ppElemFailBottom, out NetOffice.MSHTMLApi.IHTMLElement ppElemFailTop);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSegmentList">NetOffice.MSHTMLApi.ISegmentList pSegmentList</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SaveSegmentsToClipboard(NetOffice.MSHTMLApi.ISegmentList pSegmentList, Int32 dwFlags);

		#endregion
	}
}
