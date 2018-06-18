using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLNamespace_onreadystatechangeEventHandler(NetOffice.MSHTMLApi.IHTMLEventObj pEvtObj);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLNamespace 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSHTMLApi.EventContracts.HTMLNamespaceEvents))]
	[TypeId("3050F6BC-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface HTMLNamespace : DispHTMLNamespace, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		event HTMLNamespace_onreadystatechangeEventHandler onreadystatechangeEvent;

		#endregion
	}
}
