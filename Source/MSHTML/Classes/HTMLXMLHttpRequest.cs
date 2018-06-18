using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLXMLHttpRequest_ontimeoutEventHandler();
	public delegate void HTMLXMLHttpRequest_onreadystatechangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLXMLHttpRequest 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLXMLHttpRequestEvents))]
	[TypeId("3051040B-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLXMLHttpRequest : DispHTMLXMLHttpRequest, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLXMLHttpRequest_ontimeoutEventHandler ontimeoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLXMLHttpRequest_onreadystatechangeEventHandler onreadystatechangeEvent;

		#endregion
	}
}
