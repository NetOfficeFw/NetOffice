using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface HTMLWindowEvents2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F625-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface HTMLWindowEvents2 : DispHTMLWindow2
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onload(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onunload(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new bool onhelp(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onfocus(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onblur(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="description">string description</param>
		/// <param name="url">string url</param>
		/// <param name="line">Int32 line</param>
		[SupportByVersion("MSHTML", 4)]
        new void onerror(string description, string url, Int32 line);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onresize(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onscroll(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onbeforeunload(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onbeforeprint(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onafterprint(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		#endregion
	}
}
