using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ComboBox_BeforeDragOverEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.Enums.fmDragState dragState, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void ComboBox_BeforeDropOrPasteEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Enums.fmAction action, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void ComboBox_ChangeEventHandler();
	public delegate void ComboBox_ClickEventHandler();
	public delegate void ComboBox_DblClickEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void ComboBox_DropButtonClickEventHandler();
	public delegate void ComboBox_ErrorEventHandler(Int16 number, NetOffice.MSFormsApi.ReturnString description, Int32 sCode, string source, string helpFile, Int32 helpContext, NetOffice.MSFormsApi.ReturnBoolean cancelDisplay);
	public delegate void ComboBox_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void ComboBox_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger keyAscii);
	public delegate void ComboBox_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void ComboBox_MouseDownEventHandler(Int16 button, Int16 shift, Single x, Single y);
	public delegate void ComboBox_MouseMoveEventHandler(Int16 button, Int16 shift, Single x, Single y);
	public delegate void ComboBox_MouseUpEventHandler(Int16 button, Int16 shift, Single x, Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ComboBox 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.MdcComboEvents))]
	[TypeId("8BD21D30-EC42-11CE-9E0D-00AA006002F3")]
    public interface ComboBox : IMdcCombo, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_BeforeDragOverEventHandler BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_DropButtonClickEventHandler DropButtonClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ComboBox_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
