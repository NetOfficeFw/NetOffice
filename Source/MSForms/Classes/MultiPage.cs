using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	#region Delegates

	#pragma warning disable
	public delegate void MultiPage_AddControlEventHandler(Int32 index, NetOffice.MSFormsApi.Control control);
	public delegate void MultiPage_BeforeDragOverEventHandler(Int32 index, NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Control control, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.Enums.fmDragState state, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void MultiPage_BeforeDropOrPasteEventHandler(Int32 index, NetOffice.MSFormsApi.ReturnBoolean cancel, NetOffice.MSFormsApi.Control control, NetOffice.MSFormsApi.Enums.fmAction action, NetOffice.MSFormsApi.DataObject data, Single x, Single y, NetOffice.MSFormsApi.ReturnEffect effect, Int16 shift);
	public delegate void MultiPage_ChangeEventHandler();
	public delegate void MultiPage_ClickEventHandler(Int32 index);
	public delegate void MultiPage_DblClickEventHandler(Int32 index, NetOffice.MSFormsApi.ReturnBoolean cancel);
	public delegate void MultiPage_ErrorEventHandler(Int32 index, Int16 number, NetOffice.MSFormsApi.ReturnString description, Int32 sCode, string source, string helpFile, Int32 helpContext, NetOffice.MSFormsApi.ReturnBoolean cancelDisplay);
	public delegate void MultiPage_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void MultiPage_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger keyAscii);
	public delegate void MultiPage_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger keyCode, Int16 shift);
	public delegate void MultiPage_LayoutEventHandler(Int32 index);
	public delegate void MultiPage_MouseDownEventHandler(Int32 index, Int16 button, Int16 shift, Single x, Single y);
	public delegate void MultiPage_MouseMoveEventHandler(Int32 Index, Int16 button, Int16 shift, Single x, Single y);
	public delegate void MultiPage_MouseUpEventHandler(Int32 index, Int16 button, Int16 shift, Single x, Single y);
	public delegate void MultiPage_RemoveControlEventHandler(Int32 index, NetOffice.MSFormsApi.Control control);
	public delegate void MultiPage_ScrollEventHandler(Int32 index, NetOffice.MSFormsApi.Enums.fmScrollAction actionX, NetOffice.MSFormsApi.Enums.fmScrollAction actionY, Single requestDx, Single requestDy, NetOffice.MSFormsApi.ReturnSingle actualDx, NetOffice.MSFormsApi.ReturnSingle actualDy);
	public delegate void MultiPage_ZoomEventHandler(Int32 index, ref Int16 percent);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass MultiPage 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.MultiPageEvents))]
	[TypeId("46E31370-3F7A-11CE-BED6-00AA00611080")]
    public interface MultiPage : IMultiPage, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_AddControlEventHandler AddControlEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_BeforeDragOverEventHandler BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_ErrorEventHandler ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_LayoutEventHandler LayoutEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_RemoveControlEventHandler RemoveControlEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_ScrollEventHandler ScrollEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		event MultiPage_ZoomEventHandler ZoomEvent;

		#endregion
	}
}
