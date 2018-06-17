using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void HTMLText_ClickEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass HTMLText 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.WHTMLControlEvents5))]
	[TypeId("5512D11A-5CC6-11CF-8D67-00AA00BDCE1D")]
    public interface HTMLText : IWHTMLText, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event HTMLText_ClickEventHandler ClickEvent;

		#endregion
	}
}
