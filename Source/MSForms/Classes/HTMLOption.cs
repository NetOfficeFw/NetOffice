using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLOption_ClickEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLOption 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.WHTMLControlEvents4))]
	[TypeId("5512D118-5CC6-11CF-8D67-00AA00BDCE1D")]
    public interface HTMLOption : IWHTMLOption, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event HTMLOption_ClickEventHandler ClickEvent;

		#endregion
	}
}
