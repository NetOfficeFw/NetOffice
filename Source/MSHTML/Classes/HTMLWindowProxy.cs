using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLWindowProxy_onloadEventHandler();
	public delegate void HTMLWindowProxy_onunloadEventHandler();
	public delegate void HTMLWindowProxy_onhelpEventHandler();
	public delegate void HTMLWindowProxy_onfocusEventHandler();
	public delegate void HTMLWindowProxy_onblurEventHandler();
	public delegate void HTMLWindowProxy_onerrorEventHandler(string description, string url, Int32 line);
	public delegate void HTMLWindowProxy_onresizeEventHandler();
	public delegate void HTMLWindowProxy_onscrollEventHandler();
	public delegate void HTMLWindowProxy_onbeforeunloadEventHandler();
	public delegate void HTMLWindowProxy_onbeforeprintEventHandler();
	public delegate void HTMLWindowProxy_onafterprintEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLWindowProxy 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLWindowEvents))]
	[TypeId("3050F391-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLWindowProxy : DispHTMLWindowProxy, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onloadEventHandler onloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onunloadEventHandler onunloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onerrorEventHandler onerrorEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onbeforeunloadEventHandler onbeforeunloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onbeforeprintEventHandler onbeforeprintEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLWindowProxy_onafterprintEventHandler onafterprintEvent;

		#endregion
	}
}
