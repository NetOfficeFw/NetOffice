using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void SpinButton_BeforeDragOverEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.Enums.fmDragState dragState, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void SpinButton_BeforeDropOrPasteEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Enums.fmAction action, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void SpinButton_ChangeEventHandler();
	public delegate void SpinButton_ErrorEventHandler(Int16 number, NetOffice.MSFormsApi.ReturnString description, Int32 sCode, string source, string helpFile, Int32 helpContext, NetOffice.MSFormsApi.ReturnBoolean cancelDisplay);
	public delegate void SpinButton_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void SpinButton_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger keyAscii);
	public delegate void SpinButton_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void SpinButton_SpinUpEventHandler();
	public delegate void SpinButton_SpinDownEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass SpinButton 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.SpinbuttonEvents))]
	[TypeId("79176FB0-B7F2-11CE-97EF-00AA006D2776")]
    public interface SpinButton : ISpinbutton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_BeforeDragOverEventHandler BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_SpinUpEventHandler SpinUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event SpinButton_SpinDownEventHandler SpinDownEvent;

		#endregion
	}
}
