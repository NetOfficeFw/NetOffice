using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OldHTMLFormElement_onhelpEventHandler();
	public delegate void OldHTMLFormElement_onclickEventHandler();
	public delegate void OldHTMLFormElement_ondblclickEventHandler();
	public delegate void OldHTMLFormElement_onkeypressEventHandler();
	public delegate void OldHTMLFormElement_onkeydownEventHandler();
	public delegate void OldHTMLFormElement_onkeyupEventHandler();
	public delegate void OldHTMLFormElement_onmouseoutEventHandler();
	public delegate void OldHTMLFormElement_onmouseoverEventHandler();
	public delegate void OldHTMLFormElement_onmousemoveEventHandler();
	public delegate void OldHTMLFormElement_onmousedownEventHandler();
	public delegate void OldHTMLFormElement_onmouseupEventHandler();
	public delegate void OldHTMLFormElement_onselectstartEventHandler();
	public delegate void OldHTMLFormElement_onfilterchangeEventHandler();
	public delegate void OldHTMLFormElement_ondragstartEventHandler();
	public delegate void OldHTMLFormElement_onbeforeupdateEventHandler();
	public delegate void OldHTMLFormElement_onafterupdateEventHandler();
	public delegate void OldHTMLFormElement_onerrorupdateEventHandler();
	public delegate void OldHTMLFormElement_onrowexitEventHandler();
	public delegate void OldHTMLFormElement_onrowenterEventHandler();
	public delegate void OldHTMLFormElement_ondatasetchangedEventHandler();
	public delegate void OldHTMLFormElement_ondataavailableEventHandler();
	public delegate void OldHTMLFormElement_ondatasetcompleteEventHandler();
	public delegate void OldHTMLFormElement_onlosecaptureEventHandler();
	public delegate void OldHTMLFormElement_onpropertychangeEventHandler();
	public delegate void OldHTMLFormElement_onscrollEventHandler();
	public delegate void OldHTMLFormElement_onfocusEventHandler();
	public delegate void OldHTMLFormElement_onblurEventHandler();
	public delegate void OldHTMLFormElement_onresizeEventHandler();
	public delegate void OldHTMLFormElement_ondragEventHandler();
	public delegate void OldHTMLFormElement_ondragendEventHandler();
	public delegate void OldHTMLFormElement_ondragenterEventHandler();
	public delegate void OldHTMLFormElement_ondragoverEventHandler();
	public delegate void OldHTMLFormElement_ondragleaveEventHandler();
	public delegate void OldHTMLFormElement_ondropEventHandler();
	public delegate void OldHTMLFormElement_onbeforecutEventHandler();
	public delegate void OldHTMLFormElement_oncutEventHandler();
	public delegate void OldHTMLFormElement_onbeforecopyEventHandler();
	public delegate void OldHTMLFormElement_oncopyEventHandler();
	public delegate void OldHTMLFormElement_onbeforepasteEventHandler();
	public delegate void OldHTMLFormElement_onpasteEventHandler();
	public delegate void OldHTMLFormElement_oncontextmenuEventHandler();
	public delegate void OldHTMLFormElement_onrowsdeleteEventHandler();
	public delegate void OldHTMLFormElement_onrowsinsertedEventHandler();
	public delegate void OldHTMLFormElement_oncellchangeEventHandler();
	public delegate void OldHTMLFormElement_onreadystatechangeEventHandler();
	public delegate void OldHTMLFormElement_onbeforeeditfocusEventHandler();
	public delegate void OldHTMLFormElement_onlayoutcompleteEventHandler();
	public delegate void OldHTMLFormElement_onpageEventHandler();
	public delegate void OldHTMLFormElement_onbeforedeactivateEventHandler();
	public delegate void OldHTMLFormElement_onbeforeactivateEventHandler();
	public delegate void OldHTMLFormElement_onmoveEventHandler();
	public delegate void OldHTMLFormElement_oncontrolselectEventHandler();
	public delegate void OldHTMLFormElement_onmovestartEventHandler();
	public delegate void OldHTMLFormElement_onmoveendEventHandler();
	public delegate void OldHTMLFormElement_onresizestartEventHandler();
	public delegate void OldHTMLFormElement_onresizeendEventHandler();
	public delegate void OldHTMLFormElement_onmouseenterEventHandler();
	public delegate void OldHTMLFormElement_onmouseleaveEventHandler();
	public delegate void OldHTMLFormElement_onmousewheelEventHandler();
	public delegate void OldHTMLFormElement_onactivateEventHandler();
	public delegate void OldHTMLFormElement_ondeactivateEventHandler();
	public delegate void OldHTMLFormElement_onfocusinEventHandler();
	public delegate void OldHTMLFormElement_onfocusoutEventHandler();
	public delegate void OldHTMLFormElement_onsubmitEventHandler();
	public delegate void OldHTMLFormElement_onresetEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OldHTMLFormElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLFormElementEvents))]
	[TypeId("0D04D285-6BEC-11CF-8B97-00AA00476DA6")]
    public interface OldHTMLFormElement : DispHTMLFormElement, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onfocusoutEventHandler onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onsubmitEventHandler onsubmitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLFormElement_onresetEventHandler onresetEvent;

		#endregion
	}
}
