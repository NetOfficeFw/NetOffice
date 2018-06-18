using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLOptionButtonElement_onhelpEventHandler();
	public delegate void HTMLOptionButtonElement_onclickEventHandler();
	public delegate void HTMLOptionButtonElement_ondblclickEventHandler();
	public delegate void HTMLOptionButtonElement_onkeypressEventHandler();
	public delegate void HTMLOptionButtonElement_onkeydownEventHandler();
	public delegate void HTMLOptionButtonElement_onkeyupEventHandler();
	public delegate void HTMLOptionButtonElement_onmouseoutEventHandler();
	public delegate void HTMLOptionButtonElement_onmouseoverEventHandler();
	public delegate void HTMLOptionButtonElement_onmousemoveEventHandler();
	public delegate void HTMLOptionButtonElement_onmousedownEventHandler();
	public delegate void HTMLOptionButtonElement_onmouseupEventHandler();
	public delegate void HTMLOptionButtonElement_onselectstartEventHandler();
	public delegate void HTMLOptionButtonElement_onfilterchangeEventHandler();
	public delegate void HTMLOptionButtonElement_ondragstartEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforeupdateEventHandler();
	public delegate void HTMLOptionButtonElement_onafterupdateEventHandler();
	public delegate void HTMLOptionButtonElement_onerrorupdateEventHandler();
	public delegate void HTMLOptionButtonElement_onrowexitEventHandler();
	public delegate void HTMLOptionButtonElement_onrowenterEventHandler();
	public delegate void HTMLOptionButtonElement_ondatasetchangedEventHandler();
	public delegate void HTMLOptionButtonElement_ondataavailableEventHandler();
	public delegate void HTMLOptionButtonElement_ondatasetcompleteEventHandler();
	public delegate void HTMLOptionButtonElement_onlosecaptureEventHandler();
	public delegate void HTMLOptionButtonElement_onpropertychangeEventHandler();
	public delegate void HTMLOptionButtonElement_onscrollEventHandler();
	public delegate void HTMLOptionButtonElement_onfocusEventHandler();
	public delegate void HTMLOptionButtonElement_onblurEventHandler();
	public delegate void HTMLOptionButtonElement_onresizeEventHandler();
	public delegate void HTMLOptionButtonElement_ondragEventHandler();
	public delegate void HTMLOptionButtonElement_ondragendEventHandler();
	public delegate void HTMLOptionButtonElement_ondragenterEventHandler();
	public delegate void HTMLOptionButtonElement_ondragoverEventHandler();
	public delegate void HTMLOptionButtonElement_ondragleaveEventHandler();
	public delegate void HTMLOptionButtonElement_ondropEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforecutEventHandler();
	public delegate void HTMLOptionButtonElement_oncutEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforecopyEventHandler();
	public delegate void HTMLOptionButtonElement_oncopyEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforepasteEventHandler();
	public delegate void HTMLOptionButtonElement_onpasteEventHandler();
	public delegate void HTMLOptionButtonElement_oncontextmenuEventHandler();
	public delegate void HTMLOptionButtonElement_onrowsdeleteEventHandler();
	public delegate void HTMLOptionButtonElement_onrowsinsertedEventHandler();
	public delegate void HTMLOptionButtonElement_oncellchangeEventHandler();
	public delegate void HTMLOptionButtonElement_onreadystatechangeEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforeeditfocusEventHandler();
	public delegate void HTMLOptionButtonElement_onlayoutcompleteEventHandler();
	public delegate void HTMLOptionButtonElement_onpageEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforedeactivateEventHandler();
	public delegate void HTMLOptionButtonElement_onbeforeactivateEventHandler();
	public delegate void HTMLOptionButtonElement_onmoveEventHandler();
	public delegate void HTMLOptionButtonElement_oncontrolselectEventHandler();
	public delegate void HTMLOptionButtonElement_onmovestartEventHandler();
	public delegate void HTMLOptionButtonElement_onmoveendEventHandler();
	public delegate void HTMLOptionButtonElement_onresizestartEventHandler();
	public delegate void HTMLOptionButtonElement_onresizeendEventHandler();
	public delegate void HTMLOptionButtonElement_onmouseenterEventHandler();
	public delegate void HTMLOptionButtonElement_onmouseleaveEventHandler();
	public delegate void HTMLOptionButtonElement_onmousewheelEventHandler();
	public delegate void HTMLOptionButtonElement_onactivateEventHandler();
	public delegate void HTMLOptionButtonElement_ondeactivateEventHandler();
	public delegate void HTMLOptionButtonElement_onfocusinEventHandler();
	public delegate void HTMLOptionButtonElement_onfocusoutEventHandler();
	public delegate void HTMLOptionButtonElement_onchangeEventHandler();
	public delegate void HTMLOptionButtonElement_onselectEventHandler();
	public delegate void HTMLOptionButtonElement_onloadEventHandler();
	public delegate void HTMLOptionButtonElement_onerrorEventHandler();
	public delegate void HTMLOptionButtonElement_onabortEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLOptionButtonElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLOptionButtonElementEvents))]
	[TypeId("3050F2BE-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLOptionButtonElement : DispIHTMLOptionButtonElement, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onfocusoutEventHandler onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onchangeEventHandler onchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onselectEventHandler onselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onloadEventHandler onloadEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onerrorEventHandler onerrorEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLOptionButtonElement_onabortEventHandler onabortEvent;

		#endregion
	}
}
