using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLFormElement_onhelpEventHandler();
	public delegate void HTMLFormElement_onclickEventHandler();
	public delegate void HTMLFormElement_ondblclickEventHandler();
	public delegate void HTMLFormElement_onkeypressEventHandler();
	public delegate void HTMLFormElement_onkeydownEventHandler();
	public delegate void HTMLFormElement_onkeyupEventHandler();
	public delegate void HTMLFormElement_onmouseoutEventHandler();
	public delegate void HTMLFormElement_onmouseoverEventHandler();
	public delegate void HTMLFormElement_onmousemoveEventHandler();
	public delegate void HTMLFormElement_onmousedownEventHandler();
	public delegate void HTMLFormElement_onmouseupEventHandler();
	public delegate void HTMLFormElement_onselectstartEventHandler();
	public delegate void HTMLFormElement_onfilterchangeEventHandler();
	public delegate void HTMLFormElement_ondragstartEventHandler();
	public delegate void HTMLFormElement_onbeforeupdateEventHandler();
	public delegate void HTMLFormElement_onafterupdateEventHandler();
	public delegate void HTMLFormElement_onerrorupdateEventHandler();
	public delegate void HTMLFormElement_onrowexitEventHandler();
	public delegate void HTMLFormElement_onrowenterEventHandler();
	public delegate void HTMLFormElement_ondatasetchangedEventHandler();
	public delegate void HTMLFormElement_ondataavailableEventHandler();
	public delegate void HTMLFormElement_ondatasetcompleteEventHandler();
	public delegate void HTMLFormElement_onlosecaptureEventHandler();
	public delegate void HTMLFormElement_onpropertychangeEventHandler();
	public delegate void HTMLFormElement_onscrollEventHandler();
	public delegate void HTMLFormElement_onfocusEventHandler();
	public delegate void HTMLFormElement_onblurEventHandler();
	public delegate void HTMLFormElement_onresizeEventHandler();
	public delegate void HTMLFormElement_ondragEventHandler();
	public delegate void HTMLFormElement_ondragendEventHandler();
	public delegate void HTMLFormElement_ondragenterEventHandler();
	public delegate void HTMLFormElement_ondragoverEventHandler();
	public delegate void HTMLFormElement_ondragleaveEventHandler();
	public delegate void HTMLFormElement_ondropEventHandler();
	public delegate void HTMLFormElement_onbeforecutEventHandler();
	public delegate void HTMLFormElement_oncutEventHandler();
	public delegate void HTMLFormElement_onbeforecopyEventHandler();
	public delegate void HTMLFormElement_oncopyEventHandler();
	public delegate void HTMLFormElement_onbeforepasteEventHandler();
	public delegate void HTMLFormElement_onpasteEventHandler();
	public delegate void HTMLFormElement_oncontextmenuEventHandler();
	public delegate void HTMLFormElement_onrowsdeleteEventHandler();
	public delegate void HTMLFormElement_onrowsinsertedEventHandler();
	public delegate void HTMLFormElement_oncellchangeEventHandler();
	public delegate void HTMLFormElement_onreadystatechangeEventHandler();
	public delegate void HTMLFormElement_onbeforeeditfocusEventHandler();
	public delegate void HTMLFormElement_onlayoutcompleteEventHandler();
	public delegate void HTMLFormElement_onpageEventHandler();
	public delegate void HTMLFormElement_onbeforedeactivateEventHandler();
	public delegate void HTMLFormElement_onbeforeactivateEventHandler();
	public delegate void HTMLFormElement_onmoveEventHandler();
	public delegate void HTMLFormElement_oncontrolselectEventHandler();
	public delegate void HTMLFormElement_onmovestartEventHandler();
	public delegate void HTMLFormElement_onmoveendEventHandler();
	public delegate void HTMLFormElement_onresizestartEventHandler();
	public delegate void HTMLFormElement_onresizeendEventHandler();
	public delegate void HTMLFormElement_onmouseenterEventHandler();
	public delegate void HTMLFormElement_onmouseleaveEventHandler();
	public delegate void HTMLFormElement_onmousewheelEventHandler();
	public delegate void HTMLFormElement_onactivateEventHandler();
	public delegate void HTMLFormElement_ondeactivateEventHandler();
	public delegate void HTMLFormElement_onfocusinEventHandler();
	public delegate void HTMLFormElement_onfocusoutEventHandler();
	public delegate void HTMLFormElement_onsubmitEventHandler();
	public delegate void HTMLFormElement_onresetEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLFormElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLFormElementEvents))]
	[TypeId("3050F251-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLFormElement : DispHTMLFormElement, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onfocusoutEventHandler onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onsubmitEventHandler onsubmitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFormElement_onresetEventHandler onresetEvent;

		#endregion
	}
}
