using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLReset_ClickEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLReset 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.WHTMLControlEvents2))]
	[TypeId("5512D114-5CC6-11CF-8D67-00AA00BDCE1D")]
    public interface HTMLReset : IWHTMLReset, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event HTMLReset_ClickEventHandler ClickEvent;

		#endregion
	}
}
