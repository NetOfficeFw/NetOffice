using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLSelect_ClickEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLSelect 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.WHTMLControlEvents9))]
	[TypeId("5512D122-5CC6-11CF-8D67-00AA00BDCE1D")]
    public interface HTMLSelect : IWHTMLSelect, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event HTMLSelect_ClickEventHandler ClickEvent;

		#endregion
	}
}
