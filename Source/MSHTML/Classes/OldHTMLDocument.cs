using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OldHTMLDocument_onhelpEventHandler();
	public delegate void OldHTMLDocument_onclickEventHandler();
	public delegate void OldHTMLDocument_ondblclickEventHandler();
	public delegate void OldHTMLDocument_onkeydownEventHandler();
	public delegate void OldHTMLDocument_onkeyupEventHandler();
	public delegate void OldHTMLDocument_onkeypressEventHandler();
	public delegate void OldHTMLDocument_onmousedownEventHandler();
	public delegate void OldHTMLDocument_onmousemoveEventHandler();
	public delegate void OldHTMLDocument_onmouseupEventHandler();
	public delegate void OldHTMLDocument_onmouseoutEventHandler();
	public delegate void OldHTMLDocument_onmouseoverEventHandler();
	public delegate void OldHTMLDocument_onreadystatechangeEventHandler();
	public delegate void OldHTMLDocument_onbeforeupdateEventHandler();
	public delegate void OldHTMLDocument_onafterupdateEventHandler();
	public delegate void OldHTMLDocument_onrowexitEventHandler();
	public delegate void OldHTMLDocument_onrowenterEventHandler();
	public delegate void OldHTMLDocument_ondragstartEventHandler();
	public delegate void OldHTMLDocument_onselectstartEventHandler();
	public delegate void OldHTMLDocument_onerrorupdateEventHandler();
	public delegate void OldHTMLDocument_oncontextmenuEventHandler();
	public delegate void OldHTMLDocument_onstopEventHandler();
	public delegate void OldHTMLDocument_onrowsdeleteEventHandler();
	public delegate void OldHTMLDocument_onrowsinsertedEventHandler();
	public delegate void OldHTMLDocument_oncellchangeEventHandler();
	public delegate void OldHTMLDocument_onpropertychangeEventHandler();
	public delegate void OldHTMLDocument_ondatasetchangedEventHandler();
	public delegate void OldHTMLDocument_ondataavailableEventHandler();
	public delegate void OldHTMLDocument_ondatasetcompleteEventHandler();
	public delegate void OldHTMLDocument_onbeforeeditfocusEventHandler();
	public delegate void OldHTMLDocument_onselectionchangeEventHandler();
	public delegate void OldHTMLDocument_oncontrolselectEventHandler();
	public delegate void OldHTMLDocument_onmousewheelEventHandler();
	public delegate void OldHTMLDocument_onfocusinEventHandler();
	public delegate void OldHTMLDocument_onfocusoutEventHandler();
	public delegate void OldHTMLDocument_onactivateEventHandler();
	public delegate void OldHTMLDocument_ondeactivateEventHandler();
	public delegate void OldHTMLDocument_onbeforeactivateEventHandler();
	public delegate void OldHTMLDocument_onbeforedeactivateEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OldHTMLDocument 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLDocumentEvents))]
	[TypeId("D48A6EC9-6A4A-11CF-94A7-444553540000")]
    public interface OldHTMLDocument : DispHTMLDocument, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onhelpEventHandler onhelpEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onmouseupEventHandler onmouseupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onmouseoutEventHandler onmouseoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onmouseoverEventHandler onmouseoverEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_ondragstartEventHandler ondragstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onselectstartEventHandler onselectstartEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_oncontextmenuEventHandler oncontextmenuEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onstopEventHandler onstopEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onpropertychangeEventHandler onpropertychangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onbeforeeditfocusEventHandler onbeforeeditfocusEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onselectionchangeEventHandler onselectionchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_oncontrolselectEventHandler oncontrolselectEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onmousewheelEventHandler onmousewheelEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onfocusinEventHandler onfocusinEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onfocusoutEventHandler onfocusoutEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onactivateEventHandler onactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_ondeactivateEventHandler ondeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onbeforeactivateEventHandler onbeforeactivateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event OldHTMLDocument_onbeforedeactivateEventHandler onbeforedeactivateEvent;

		#endregion
	}
}
