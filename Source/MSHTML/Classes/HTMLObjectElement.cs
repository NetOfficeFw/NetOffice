using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLObjectElement_onbeforeupdateEventHandler();
	public delegate void HTMLObjectElement_onafterupdateEventHandler();
	public delegate void HTMLObjectElement_onerrorupdateEventHandler();
	public delegate void HTMLObjectElement_onrowexitEventHandler();
	public delegate void HTMLObjectElement_onrowenterEventHandler();
	public delegate void HTMLObjectElement_ondatasetchangedEventHandler();
	public delegate void HTMLObjectElement_ondataavailableEventHandler();
	public delegate void HTMLObjectElement_ondatasetcompleteEventHandler();
	public delegate void HTMLObjectElement_onerrorEventHandler();
	public delegate void HTMLObjectElement_onrowsdeleteEventHandler();
	public delegate void HTMLObjectElement_onrowsinsertedEventHandler();
	public delegate void HTMLObjectElement_oncellchangeEventHandler();
	public delegate void HTMLObjectElement_onreadystatechangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLObjectElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLObjectElementEvents))]
	[TypeId("3050F24E-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLObjectElement : DispHTMLObjectElement, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onbeforeupdateEventHandler onbeforeupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onafterupdateEventHandler onafterupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onerrorupdateEventHandler onerrorupdateEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onrowexitEventHandler onrowexitEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onrowenterEventHandler onrowenterEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_ondatasetchangedEventHandler ondatasetchangedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_ondataavailableEventHandler ondataavailableEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_ondatasetcompleteEventHandler ondatasetcompleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onerrorEventHandler onerrorEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onrowsdeleteEventHandler onrowsdeleteEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onrowsinsertedEventHandler onrowsinsertedEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_oncellchangeEventHandler oncellchangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLObjectElement_onreadystatechangeEventHandler onreadystatechangeEvent;

		#endregion
	}
}
