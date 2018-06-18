using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLEditDesigner 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F662-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLEditDesigner : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 PreHandleEvent(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 PostHandleEvent(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 TranslateAccelerator(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 PostEditorEventNotify(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj);

		#endregion
	}
}
