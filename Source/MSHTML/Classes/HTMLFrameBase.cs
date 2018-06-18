using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLFrameBase_onhelpEventHandler();
	public delegate void HTMLFrameBase_onclickEventHandler();
	public delegate void HTMLFrameBase_ondblclickEventHandler();
	public delegate void HTMLFrameBase_onkeypressEventHandler();
	public delegate void HTMLFrameBase_onkeydownEventHandler();
	public delegate void HTMLFrameBase_onkeyupEventHandler();
	public delegate void HTMLFrameBase_onmouseoutEventHandler();
	public delegate void HTMLFrameBase_onmouseoverEventHandler();
	public delegate void HTMLFrameBase_onmousemoveEventHandler();
	public delegate void HTMLFrameBase_onmousedownEventHandler();
	public delegate void HTMLFrameBase_onmouseupEventHandler();
	public delegate void HTMLFrameBase_onselectstartEventHandler();
	public delegate void HTMLFrameBase_onfilterchangeEventHandler();
	public delegate void HTMLFrameBase_ondragstartEventHandler();
	public delegate void HTMLFrameBase_onbeforeupdateEventHandler();
	public delegate void HTMLFrameBase_onafterupdateEventHandler();
	public delegate void HTMLFrameBase_onerrorupdateEventHandler();
	public delegate void HTMLFrameBase_onrowexitEventHandler();
	public delegate void HTMLFrameBase_onrowenterEventHandler();
	public delegate void HTMLFrameBase_ondatasetchangedEventHandler();
	public delegate void HTMLFrameBase_ondataavailableEventHandler();
	public delegate void HTMLFrameBase_ondatasetcompleteEventHandler();
	public delegate void HTMLFrameBase_onlosecaptureEventHandler();
	public delegate void HTMLFrameBase_onpropertychangeEventHandler();
	public delegate void HTMLFrameBase_onscrollEventHandler();
	public delegate void HTMLFrameBase_onfocusEventHandler();
	public delegate void HTMLFrameBase_onblurEventHandler();
	public delegate void HTMLFrameBase_onresizeEventHandler();
	public delegate void HTMLFrameBase_ondragEventHandler();
	public delegate void HTMLFrameBase_ondragendEventHandler();
	public delegate void HTMLFrameBase_ondragenterEventHandler();
	public delegate void HTMLFrameBase_ondragoverEventHandler();
	public delegate void HTMLFrameBase_ondragleaveEventHandler();
	public delegate void HTMLFrameBase_ondropEventHandler();
	public delegate void HTMLFrameBase_onbeforecutEventHandler();
	public delegate void HTMLFrameBase_oncutEventHandler();
	public delegate void HTMLFrameBase_onbeforecopyEventHandler();
	public delegate void HTMLFrameBase_oncopyEventHandler();
	public delegate void HTMLFrameBase_onbeforepasteEventHandler();
	public delegate void HTMLFrameBase_onpasteEventHandler();
	public delegate void HTMLFrameBase_oncontextmenuEventHandler();
	public delegate void HTMLFrameBase_onrowsdeleteEventHandler();
	public delegate void HTMLFrameBase_onrowsinsertedEventHandler();
	public delegate void HTMLFrameBase_oncellchangeEventHandler();
	public delegate void HTMLFrameBase_onreadystatechangeEventHandler();
	public delegate void HTMLFrameBase_onbeforeeditfocusEventHandler();
	public delegate void HTMLFrameBase_onlayoutcompleteEventHandler();
	public delegate void HTMLFrameBase_onpageEventHandler();
	public delegate void HTMLFrameBase_onbeforedeactivateEventHandler();
	public delegate void HTMLFrameBase_onbeforeactivateEventHandler();
	public delegate void HTMLFrameBase_onmoveEventHandler();
	public delegate void HTMLFrameBase_oncontrolselectEventHandler();
	public delegate void HTMLFrameBase_onmovestartEventHandler();
	public delegate void HTMLFrameBase_onmoveendEventHandler();
	public delegate void HTMLFrameBase_onresizestartEventHandler();
	public delegate void HTMLFrameBase_onresizeendEventHandler();
	public delegate void HTMLFrameBase_onmouseenterEventHandler();
	public delegate void HTMLFrameBase_onmouseleaveEventHandler();
	public delegate void HTMLFrameBase_onmousewheelEventHandler();
	public delegate void HTMLFrameBase_onactivateEventHandler();
	public delegate void HTMLFrameBase_ondeactivateEventHandler();
	public delegate void HTMLFrameBase_onfocusinEventHandler();
	public delegate void HTMLFrameBase_onfocusoutEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLFrameBase 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLControlElementEvents))]
	[TypeId("3050F312-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLFrameBase : DispHTMLFrameBase, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onfilterchangeEventHandler onfilterchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onlosecaptureEventHandler onlosecaptureEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onscrollEventHandler onscrollEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onfocusEventHandler onfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onblurEventHandler onblurEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onresizeEventHandler onresizeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondragEventHandler ondragEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondragendEventHandler ondragendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondragenterEventHandler ondragenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondragoverEventHandler ondragoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondragleaveEventHandler ondragleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondropEventHandler ondropEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforecutEventHandler onbeforecutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_oncutEventHandler oncutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforecopyEventHandler onbeforecopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_oncopyEventHandler oncopyEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforepasteEventHandler onbeforepasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onpasteEventHandler onpasteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onlayoutcompleteEventHandler onlayoutcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onpageEventHandler onpageEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmoveEventHandler onmoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmovestartEventHandler onmovestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmoveendEventHandler onmoveendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onresizestartEventHandler onresizestartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onresizeendEventHandler onresizeendEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmouseenterEventHandler onmouseenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmouseleaveEventHandler onmouseleaveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLFrameBase_onfocusoutEventHandler onfocusoutEvent;

		#endregion
	}
}
