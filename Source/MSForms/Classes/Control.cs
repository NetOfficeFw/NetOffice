using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Control_EnterEventHandler();
	public delegate void Control_ExitEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void Control_BeforeUpdateEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void Control_AfterUpdateEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Control 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ControlEvents))]
	[TypeId("909E0AE0-16DC-11CE-9E98-00AA00574A4F")]
    public interface Control : IControl, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event Control_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event Control_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event Control_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event Control_AfterUpdateEventHandler AfterUpdateEvent;

		#endregion
	}
}
