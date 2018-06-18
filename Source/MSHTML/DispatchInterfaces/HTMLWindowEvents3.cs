using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface HTMLWindowEvents3 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F5A1-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface HTMLWindowEvents3 : HTMLWindowEvents2
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onhashchange(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onmessage(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		#endregion
	}
}
