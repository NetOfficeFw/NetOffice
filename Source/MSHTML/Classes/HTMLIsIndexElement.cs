using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLIsIndexElement_onhelpEventHandler();
	public delegate void HTMLIsIndexElement_onclickEventHandler();
	public delegate void HTMLIsIndexElement_ondblclickEventHandler();
	public delegate void HTMLIsIndexElement_onkeypressEventHandler();
	public delegate void HTMLIsIndexElement_onkeydownEventHandler();
	public delegate void HTMLIsIndexElement_onkeyupEventHandler();
	public delegate void HTMLIsIndexElement_onmouseoutEventHandler();
	public delegate void HTMLIsIndexElement_onmouseoverEventHandler();
	public delegate void HTMLIsIndexElement_onmousemoveEventHandler();
	public delegate void HTMLIsIndexElement_onmousedownEventHandler();
	public delegate void HTMLIsIndexElement_onmouseupEventHandler();
	public delegate void HTMLIsIndexElement_onselectstartEventHandler();
	public delegate void HTMLIsIndexElement_onfilterchangeEventHandler();
	public delegate void HTMLIsIndexElement_ondragstartEventHandler();
	public delegate void HTMLIsIndexElement_onbeforeupdateEventHandler();
	public delegate void HTMLIsIndexElement_onafterupdateEventHandler();
	public delegate void HTMLIsIndexElement_onerrorupdateEventHandler();
	public delegate void HTMLIsIndexElement_onrowexitEventHandler();
	public delegate void HTMLIsIndexElement_onrowenterEventHandler();
	public delegate void HTMLIsIndexElement_ondatasetchangedEventHandler();
	public delegate void HTMLIsIndexElement_ondataavailableEventHandler();
	public delegate void HTMLIsIndexElement_ondatasetcompleteEventHandler();
	public delegate void HTMLIsIndexElement_onlosecaptureEventHandler();
	public delegate void HTMLIsIndexElement_onpropertychangeEventHandler();
	public delegate void HTMLIsIndexElement_onscrollEventHandler();
	public delegate void HTMLIsIndexElement_onfocusEventHandler();
	public delegate void HTMLIsIndexElement_onblurEventHandler();
	public delegate void HTMLIsIndexElement_onresizeEventHandler();
	public delegate void HTMLIsIndexElement_ondragEventHandler();
	public delegate void HTMLIsIndexElement_ondragendEventHandler();
	public delegate void HTMLIsIndexElement_ondragenterEventHandler();
	public delegate void HTMLIsIndexElement_ondragoverEventHandler();
	public delegate void HTMLIsIndexElement_ondragleaveEventHandler();
	public delegate void HTMLIsIndexElement_ondropEventHandler();
	public delegate void HTMLIsIndexElement_onbeforecutEventHandler();
	public delegate void HTMLIsIndexElement_oncutEventHandler();
	public delegate void HTMLIsIndexElement_onbeforecopyEventHandler();
	public delegate void HTMLIsIndexElement_oncopyEventHandler();
	public delegate void HTMLIsIndexElement_onbeforepasteEventHandler();
	public delegate void HTMLIsIndexElement_onpasteEventHandler();
	public delegate void HTMLIsIndexElement_oncontextmenuEventHandler();
	public delegate void HTMLIsIndexElement_onrowsdeleteEventHandler();
	public delegate void HTMLIsIndexElement_onrowsinsertedEventHandler();
	public delegate void HTMLIsIndexElement_oncellchangeEventHandler();
	public delegate void HTMLIsIndexElement_onreadystatechangeEventHandler();
	public delegate void HTMLIsIndexElement_onbeforeeditfocusEventHandler();
	public delegate void HTMLIsIndexElement_onlayoutcompleteEventHandler();
	public delegate void HTMLIsIndexElement_onpageEventHandler();
	public delegate void HTMLIsIndexElement_onbeforedeactivateEventHandler();
	public delegate void HTMLIsIndexElement_onbeforeactivateEventHandler();
	public delegate void HTMLIsIndexElement_onmoveEventHandler();
	public delegate void HTMLIsIndexElement_oncontrolselectEventHandler();
	public delegate void HTMLIsIndexElement_onmovestartEventHandler();
	public delegate void HTMLIsIndexElement_onmoveendEventHandler();
	public delegate void HTMLIsIndexElement_onresizestartEventHandler();
	public delegate void HTMLIsIndexElement_onresizeendEventHandler();
	public delegate void HTMLIsIndexElement_onmouseenterEventHandler();
	public delegate void HTMLIsIndexElement_onmouseleaveEventHandler();
	public delegate void HTMLIsIndexElement_onmousewheelEventHandler();
	public delegate void HTMLIsIndexElement_onactivateEventHandler();
	public delegate void HTMLIsIndexElement_ondeactivateEventHandler();
	public delegate void HTMLIsIndexElement_onfocusinEventHandler();
	public delegate void HTMLIsIndexElement_onfocusoutEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLIsIndexElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLElementEvents))]
	[TypeId("3050F278-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLIsIndexElement : DispHTMLIsIndexElement, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLIsIndexElement_onfocusoutEventHandler onfocusoutEvent;

		#endregion
	}
}
