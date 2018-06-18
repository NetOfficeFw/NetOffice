using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Scriptlet_onscriptleteventEventHandler(string name, object eventData);
	public delegate void Scriptlet_onreadystatechangeEventHandler();
	public delegate void Scriptlet_onclickEventHandler();
	public delegate void Scriptlet_ondblclickEventHandler();
	public delegate void Scriptlet_onkeydownEventHandler();
	public delegate void Scriptlet_onkeyupEventHandler();
	public delegate void Scriptlet_onkeypressEventHandler();
	public delegate void Scriptlet_onmousedownEventHandler();
	public delegate void Scriptlet_onmousemoveEventHandler();
	public delegate void Scriptlet_onmouseupEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Scriptlet 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.DWebBridgeEvents))]
	[TypeId("AE24FDAE-03C6-11D1-8B76-0080C744F389")]
    public interface Scriptlet : IWebBridge, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onscriptleteventEventHandler onscriptleteventEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onreadystatechangeEventHandler onreadystatechangeEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onclickEventHandler onclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_ondblclickEventHandler ondblclickEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onkeydownEventHandler onkeydownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onkeyupEventHandler onkeyupEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onkeypressEventHandler onkeypressEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onmousedownEventHandler onmousedownEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onmousemoveEventHandler onmousemoveEvent;

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event Scriptlet_onmouseupEventHandler onmouseupEvent;

		#endregion
	}
}
