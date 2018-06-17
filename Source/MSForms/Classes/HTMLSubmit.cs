using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLSubmit_ClickEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLSubmit 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.WHTMLControlEvents))]
	[TypeId("5512D110-5CC6-11CF-8D67-00AA00BDCE1D")]
    public interface HTMLSubmit : IWHTMLSubmitButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event HTMLSubmit_ClickEventHandler ClickEvent;

		#endregion
	}
}
