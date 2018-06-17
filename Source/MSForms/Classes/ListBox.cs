using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ListBox_BeforeDragOverEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.Enums.fmDragState dragState, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void ListBox_BeforeDropOrPasteEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Enums.fmAction action, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void ListBox_ChangeEventHandler();
	public delegate void ListBox_ClickEventHandler();
	public delegate void ListBox_DblClickEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void ListBox_ErrorEventHandler(Int16 number, NetOffice.MSFormsApi.ReturnString description, Int32 sCode, string source, string helpFile, Int32 helpContext, NetOffice.MSFormsApi.ReturnBoolean cancelDisplay);
	public delegate void ListBox_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void ListBox_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger keyAscii);
	public delegate void ListBox_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void ListBox_MouseDownEventHandler(Int16 button, Int16 shift, Single x, Single y);
	public delegate void ListBox_MouseMoveEventHandler(Int16 button, Int16 shift, Single x, Single y);
	public delegate void ListBox_MouseUpEventHandler(Int16 button, Int16 shift, Single x, Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ListBox 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.MdcListEvents))]
	[TypeId("8BD21D20-EC42-11CE-9E0D-00AA006002F3")]
    public interface ListBox : IMdcList, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_BeforeDragOverEventHandler BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event ListBox_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
