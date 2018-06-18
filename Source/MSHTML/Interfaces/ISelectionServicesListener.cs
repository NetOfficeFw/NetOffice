using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISelectionServicesListener 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F699-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface ISelectionServicesListener : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 BeginSelectionUndo();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 EndSelectionUndo();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElementStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementStart</param>
		/// <param name="pIElementEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd</param>
		/// <param name="pIElementContentStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart</param>
		/// <param name="pIElementContentEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 OnSelectedElementExit(NetOffice.MSHTMLApi.IMarkupPointer pIElementStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 OnChangeType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pTypeDetail">string pTypeDetail</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetTypeDetail(out string pTypeDetail);

		#endregion
	}
}
