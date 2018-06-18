using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLFrameSetSite_onhelpEventHandler();
	public delegate void HTMLFrameSetSite_onclickEventHandler();
	public delegate void HTMLFrameSetSite_ondblclickEventHandler();
	public delegate void HTMLFrameSetSite_onkeypressEventHandler();
	public delegate void HTMLFrameSetSite_onkeydownEventHandler();
	public delegate void HTMLFrameSetSite_onkeyupEventHandler();
	public delegate void HTMLFrameSetSite_onmouseoutEventHandler();
	public delegate void HTMLFrameSetSite_onmouseoverEventHandler();
	public delegate void HTMLFrameSetSite_onmousemoveEventHandler();
	public delegate void HTMLFrameSetSite_onmousedownEventHandler();
	public delegate void HTMLFrameSetSite_onmouseupEventHandler();
	public delegate void HTMLFrameSetSite_onselectstartEventHandler();
	public delegate void HTMLFrameSetSite_onfilterchangeEventHandler();
	public delegate void HTMLFrameSetSite_ondragstartEventHandler();
	public delegate void HTMLFrameSetSite_onbeforeupdateEventHandler();
	public delegate void HTMLFrameSetSite_onafterupdateEventHandler();
	public delegate void HTMLFrameSetSite_onerrorupdateEventHandler();
	public delegate void HTMLFrameSetSite_onrowexitEventHandler();
	public delegate void HTMLFrameSetSite_onrowenterEventHandler();
	public delegate void HTMLFrameSetSite_ondatasetchangedEventHandler();
	public delegate void HTMLFrameSetSite_ondataavailableEventHandler();
	public delegate void HTMLFrameSetSite_ondatasetcompleteEventHandler();
	public delegate void HTMLFrameSetSite_onlosecaptureEventHandler();
	public delegate void HTMLFrameSetSite_onpropertychangeEventHandler();
	public delegate void HTMLFrameSetSite_onscrollEventHandler();
	public delegate void HTMLFrameSetSite_onfocusEventHandler();
	public delegate void HTMLFrameSetSite_onblurEventHandler();
	public delegate void HTMLFrameSetSite_onresizeEventHandler();
	public delegate void HTMLFrameSetSite_ondragEventHandler();
	public delegate void HTMLFrameSetSite_ondragendEventHandler();
	public delegate void HTMLFrameSetSite_ondragenterEventHandler();
	public delegate void HTMLFrameSetSite_ondragoverEventHandler();
	public delegate void HTMLFrameSetSite_ondragleaveEventHandler();
	public delegate void HTMLFrameSetSite_ondropEventHandler();
	public delegate void HTMLFrameSetSite_onbeforecutEventHandler();
	public delegate void HTMLFrameSetSite_oncutEventHandler();
	public delegate void HTMLFrameSetSite_onbeforecopyEventHandler();
	public delegate void HTMLFrameSetSite_oncopyEventHandler();
	public delegate void HTMLFrameSetSite_onbeforepasteEventHandler();
	public delegate void HTMLFrameSetSite_onpasteEventHandler();
	public delegate void HTMLFrameSetSite_oncontextmenuEventHandler();
	public delegate void HTMLFrameSetSite_onrowsdeleteEventHandler();
	public delegate void HTMLFrameSetSite_onrowsinsertedEventHandler();
	public delegate void HTMLFrameSetSite_oncellchangeEventHandler();
	public delegate void HTMLFrameSetSite_onreadystatechangeEventHandler();
	public delegate void HTMLFrameSetSite_onbeforeeditfocusEventHandler();
	public delegate void HTMLFrameSetSite_onlayoutcompleteEventHandler();
	public delegate void HTMLFrameSetSite_onpageEventHandler();
	public delegate void HTMLFrameSetSite_onbeforedeactivateEventHandler();
	public delegate void HTMLFrameSetSite_onbeforeactivateEventHandler();
	public delegate void HTMLFrameSetSite_onmoveEventHandler();
	public delegate void HTMLFrameSetSite_oncontrolselectEventHandler();
	public delegate void HTMLFrameSetSite_onmovestartEventHandler();
	public delegate void HTMLFrameSetSite_onmoveendEventHandler();
	public delegate void HTMLFrameSetSite_onresizestartEventHandler();
	public delegate void HTMLFrameSetSite_onresizeendEventHandler();
	public delegate void HTMLFrameSetSite_onmouseenterEventHandler();
	public delegate void HTMLFrameSetSite_onmouseleaveEventHandler();
	public delegate void HTMLFrameSetSite_onmousewheelEventHandler();
	public delegate void HTMLFrameSetSite_onactivateEventHandler();
	public delegate void HTMLFrameSetSite_ondeactivateEventHandler();
	public delegate void HTMLFrameSetSite_onfocusinEventHandler();
	public delegate void HTMLFrameSetSite_onfocusoutEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLFrameSetSite 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLControlElementEvents))]
	[TypeId("3050F31A-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLFrameSetSite : DispHTMLFrameSetSite, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameSetSite_onfocusoutEventHandler onfocusoutEvent;

		#endregion
	}
}
