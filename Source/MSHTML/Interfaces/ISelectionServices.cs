using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISelectionServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F684-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface ISelectionServices : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetSelectionType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIContainer">NetOffice.MSHTMLApi.IMarkupContainer ppIContainer</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetMarkupContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppIContainer);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStart">NetOffice.MSHTMLApi.IMarkupPointer pIStart</param>
		/// <param name="pIEnd">NetOffice.MSHTMLApi.IMarkupPointer pIEnd</param>
		/// <param name="ppISegmentAdded">NetOffice.MSHTMLApi.ISegment ppISegmentAdded</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AddSegment(NetOffice.MSHTMLApi.IMarkupPointer pIStart, NetOffice.MSHTMLApi.IMarkupPointer pIEnd, out NetOffice.MSHTMLApi.ISegment ppISegmentAdded);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="ppISegmentAdded">NetOffice.MSHTMLApi.IElementSegment ppISegmentAdded</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AddElementSegment(NetOffice.MSHTMLApi.IHTMLElement pIElement, out NetOffice.MSHTMLApi.IElementSegment ppISegmentAdded);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.ISegment pISegment</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RemoveSegment(NetOffice.MSHTMLApi.ISegment pISegment);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppISelectionServicesListener">NetOffice.MSHTMLApi.ISelectionServicesListener ppISelectionServicesListener</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSelectionServicesListener(out NetOffice.MSHTMLApi.ISelectionServicesListener ppISelectionServicesListener);

		#endregion
	}
}
