using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OptionButton_BeforeDragOverEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.Enums.fmDragState dragState, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void OptionButton_BeforeDropOrPasteEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Enums.fmAction action, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void OptionButton_ChangeEventHandler();
	public delegate void OptionButton_ClickEventHandler();
	public delegate void OptionButton_DblClickEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void OptionButton_ErrorEventHandler(Int16 number, NetOffice.MSFormsApi.ReturnString description, Int32 sCode, string source, string helpFile, Int32 helpContext, NetOffice.MSFormsApi.ReturnBoolean cancelDisplay);
	public delegate void OptionButton_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void OptionButton_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode);
	public delegate void OptionButton_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void OptionButton_MouseDownEventHandler(Int16 button, Int16 shift, Single x, Single y);
	public delegate void OptionButton_MouseMoveEventHandler(Int16 button, Int16 shift, Single x, Single y);
	public delegate void OptionButton_MouseUpEventHandler(Int16 button, Int16 shift, Single x, Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OptionButton 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.MdcOptionButtonEvents))]
	[TypeId("8BD21D50-EC42-11CE-9E0D-00AA006002F3")]
    public interface OptionButton : IMdcOptionButton, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_BeforeDragOverEventHandler BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event OptionButton_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
