using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLWindow2_onloadEventHandler();
	public delegate void HTMLWindow2_onunloadEventHandler();
	public delegate void HTMLWindow2_onhelpEventHandler();
	public delegate void HTMLWindow2_onfocusEventHandler();
	public delegate void HTMLWindow2_onblurEventHandler();
	public delegate void HTMLWindow2_onerrorEventHandler(string description, string url, Int32 line);
	public delegate void HTMLWindow2_onresizeEventHandler();
	public delegate void HTMLWindow2_onscrollEventHandler();
	public delegate void HTMLWindow2_onbeforeunloadEventHandler();
	public delegate void HTMLWindow2_onbeforeprintEventHandler();
	public delegate void HTMLWindow2_onafterprintEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLWindow2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLWindowEvents))]
	[TypeId("D48A6EC6-6A4A-11CF-94A7-444553540000")]
    public interface HTMLWindow2 : DispHTMLWindow2, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onloadEventHandler onloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onunloadEventHandler onunloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onerrorEventHandler onerrorEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onbeforeunloadEventHandler onbeforeunloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onbeforeprintEventHandler onbeforeprintEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindow2_onafterprintEventHandler onafterprintEvent;

		#endregion
	}
}
