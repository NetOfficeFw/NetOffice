using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface HTMLDocumentEvents3 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F5A0-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface HTMLDocumentEvents3 : HTMLDocumentEvents2
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onstorage(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pEvtObj">NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj</param>
		[SupportByVersion("MSHTML", 4)]
        new void onstoragecommit(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);

		#endregion
	}
}
