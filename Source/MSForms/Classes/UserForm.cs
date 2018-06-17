using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void UserForm_AddControlEventHandler(NetOffice.MSFormsApi.Control control);
	public delegate void UserForm_BeforeDragOverEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Control control, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.Enums.fmDragState state, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void UserForm_BeforeDropOrPasteEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Control control, NetOffice.MSFormsApi.Enums.fmAction action, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void UserForm_ClickEventHandler();
	public delegate void UserForm_DblClickEventHandler(NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void UserForm_ErrorEventHandler(Int16 number, NetOffice.MSFormsApi.ReturnString description, Int32 sCode, string source, string helpFile, Int32 helpContext, NetOffice.MSFormsApi.ReturnBoolean cancelDisplay);
	public delegate void UserForm_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void UserForm_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger keyAscii);
	public delegate void UserForm_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 Shift);
	public delegate void UserForm_LayoutEventHandler();
	public delegate void UserForm_MouseDownEventHandler(Int16 button, Int16 shift, Single X, Single Y);
	public delegate void UserForm_MouseMoveEventHandler(Int16 button, Int16 shift, Single X, Single Y);
	public delegate void UserForm_MouseUpEventHandler(Int16 button, Int16 shift, Single X, Single Y);
	public delegate void UserForm_RemoveControlEventHandler(NetOffice.MSFormsApi.Control control);
	public delegate void UserForm_ScrollEventHandler(NetOffice.MSFormsApi.Enums.fmScrollAction actionX, NetOffice.MSFormsApi.Enums.fmScrollAction actionY, Single requestDx, Single requestDy, NetOffice.MSFormsApi.ReturnSingle actualDx, NetOffice.MSFormsApi.ReturnSingle actualDy);
	public delegate void UserForm_ZoomEventHandler(ref Int16 percent);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass UserForm 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.FormEvents))]
	[TypeId("C62A69F0-16DC-11CE-9E98-00AA00574A4F")]
    public interface UserForm : _UserForm, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_AddControlEventHandler AddControlEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_BeforeDragOverEventHandler BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_LayoutEventHandler LayoutEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_RemoveControlEventHandler RemoveControlEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_ScrollEventHandler ScrollEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event UserForm_ZoomEventHandler ZoomEvent;

		#endregion
	}
}
