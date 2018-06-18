using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLEditServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F663-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLEditServices : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AddDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RemoveDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIContainer">NetOffice.MSHTMLApi.IMarkupContainer pIContainer</param>
		/// <param name="ppSelSvc">NetOffice.MSHTMLApi.ISelectionServices ppSelSvc</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSelectionServices(NetOffice.MSHTMLApi.IMarkupContainer pIContainer, out NetOffice.MSHTMLApi.ISelectionServices ppSelSvc);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToSelectionAnchor(NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToSelectionEnd(NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pStart">NetOffice.MSHTMLApi.IMarkupPointer pStart</param>
		/// <param name="pEnd">NetOffice.MSHTMLApi.IMarkupPointer pEnd</param>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SelectRange(NetOffice.MSHTMLApi.IMarkupPointer pStart, NetOffice.MSHTMLApi.IMarkupPointer pEnd, NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType);

		#endregion
	}
}
