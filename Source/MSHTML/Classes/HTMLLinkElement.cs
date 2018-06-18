using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLLinkElement_onhelpEventHandler();
	public delegate void HTMLLinkElement_onclickEventHandler();
	public delegate void HTMLLinkElement_ondblclickEventHandler();
	public delegate void HTMLLinkElement_onkeypressEventHandler();
	public delegate void HTMLLinkElement_onkeydownEventHandler();
	public delegate void HTMLLinkElement_onkeyupEventHandler();
	public delegate void HTMLLinkElement_onmouseoutEventHandler();
	public delegate void HTMLLinkElement_onmouseoverEventHandler();
	public delegate void HTMLLinkElement_onmousemoveEventHandler();
	public delegate void HTMLLinkElement_onmousedownEventHandler();
	public delegate void HTMLLinkElement_onmouseupEventHandler();
	public delegate void HTMLLinkElement_onselectstartEventHandler();
	public delegate void HTMLLinkElement_onfilterchangeEventHandler();
	public delegate void HTMLLinkElement_ondragstartEventHandler();
	public delegate void HTMLLinkElement_onbeforeupdateEventHandler();
	public delegate void HTMLLinkElement_onafterupdateEventHandler();
	public delegate void HTMLLinkElement_onerrorupdateEventHandler();
	public delegate void HTMLLinkElement_onrowexitEventHandler();
	public delegate void HTMLLinkElement_onrowenterEventHandler();
	public delegate void HTMLLinkElement_ondatasetchangedEventHandler();
	public delegate void HTMLLinkElement_ondataavailableEventHandler();
	public delegate void HTMLLinkElement_ondatasetcompleteEventHandler();
	public delegate void HTMLLinkElement_onlosecaptureEventHandler();
	public delegate void HTMLLinkElement_onpropertychangeEventHandler();
	public delegate void HTMLLinkElement_onscrollEventHandler();
	public delegate void HTMLLinkElement_onfocusEventHandler();
	public delegate void HTMLLinkElement_onblurEventHandler();
	public delegate void HTMLLinkElement_onresizeEventHandler();
	public delegate void HTMLLinkElement_ondragEventHandler();
	public delegate void HTMLLinkElement_ondragendEventHandler();
	public delegate void HTMLLinkElement_ondragenterEventHandler();
	public delegate void HTMLLinkElement_ondragoverEventHandler();
	public delegate void HTMLLinkElement_ondragleaveEventHandler();
	public delegate void HTMLLinkElement_ondropEventHandler();
	public delegate void HTMLLinkElement_onbeforecutEventHandler();
	public delegate void HTMLLinkElement_oncutEventHandler();
	public delegate void HTMLLinkElement_onbeforecopyEventHandler();
	public delegate void HTMLLinkElement_oncopyEventHandler();
	public delegate void HTMLLinkElement_onbeforepasteEventHandler();
	public delegate void HTMLLinkElement_onpasteEventHandler();
	public delegate void HTMLLinkElement_oncontextmenuEventHandler();
	public delegate void HTMLLinkElement_onrowsdeleteEventHandler();
	public delegate void HTMLLinkElement_onrowsinsertedEventHandler();
	public delegate void HTMLLinkElement_oncellchangeEventHandler();
	public delegate void HTMLLinkElement_onreadystatechangeEventHandler();
	public delegate void HTMLLinkElement_onbeforeeditfocusEventHandler();
	public delegate void HTMLLinkElement_onlayoutcompleteEventHandler();
	public delegate void HTMLLinkElement_onpageEventHandler();
	public delegate void HTMLLinkElement_onbeforedeactivateEventHandler();
	public delegate void HTMLLinkElement_onbeforeactivateEventHandler();
	public delegate void HTMLLinkElement_onmoveEventHandler();
	public delegate void HTMLLinkElement_oncontrolselectEventHandler();
	public delegate void HTMLLinkElement_onmovestartEventHandler();
	public delegate void HTMLLinkElement_onmoveendEventHandler();
	public delegate void HTMLLinkElement_onresizestartEventHandler();
	public delegate void HTMLLinkElement_onresizeendEventHandler();
	public delegate void HTMLLinkElement_onmouseenterEventHandler();
	public delegate void HTMLLinkElement_onmouseleaveEventHandler();
	public delegate void HTMLLinkElement_onmousewheelEventHandler();
	public delegate void HTMLLinkElement_onactivateEventHandler();
	public delegate void HTMLLinkElement_ondeactivateEventHandler();
	public delegate void HTMLLinkElement_onfocusinEventHandler();
	public delegate void HTMLLinkElement_onfocusoutEventHandler();
	public delegate void HTMLLinkElement_onloadEventHandler();
	public delegate void HTMLLinkElement_onerrorEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLLinkElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLLinkElementEvents))]
	[TypeId("3050F277-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLLinkElement : DispHTMLLinkElement, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onfocusoutEventHandler onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onloadEventHandler onloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLLinkElement_onerrorEventHandler onerrorEvent;

		#endregion
	}
}
