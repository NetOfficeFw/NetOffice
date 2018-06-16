using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	#region Delegates

	#pragma warning disable
	public delegate void Spreadsheet_BeforeContextMenuEventHandler(Int32 x, Int32 y, NetOffice.OWC10Api.ByRef menu, NetOffice.OWC10Api.ByRef cancel);
	public delegate void Spreadsheet_BeforeKeyDownEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void Spreadsheet_BeforeKeyPressEventHandler(Int32 keyAscii, NetOffice.OWC10Api.ByRef cancel);
	public delegate void Spreadsheet_BeforeKeyUpEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void Spreadsheet_ClickEventHandler();
	public delegate void Spreadsheet_CommandEnabledEventHandler(object command, NetOffice.OWC10Api.ByRef enabled);
	public delegate void Spreadsheet_CommandCheckedEventHandler(object command, NetOffice.OWC10Api.ByRef Checked);
	public delegate void Spreadsheet_CommandTipTextEventHandler(object command, NetOffice.OWC10Api.ByRef caption);
	public delegate void Spreadsheet_CommandBeforeExecuteEventHandler(object command, NetOffice.OWC10Api.ByRef cancel);
	public delegate void Spreadsheet_CommandExecuteEventHandler(object command, bool succeeded);
	public delegate void Spreadsheet_DblClickEventHandler();
	public delegate void Spreadsheet_EndEditEventHandler(bool accept, NetOffice.OWC10Api.ByRef finalValue, NetOffice.OWC10Api.ByRef cancel, NetOffice.OWC10Api.ByRef errorDescription);
	public delegate void Spreadsheet_InitializeEventHandler();
	public delegate void Spreadsheet_KeyDownEventHandler(Int32 keyCode, Int32 shift);
	public delegate void Spreadsheet_KeyPressEventHandler(Int32 keyAscii);
	public delegate void Spreadsheet_KeyUpEventHandler(Int32 keyCode, Int32 shift);
	public delegate void Spreadsheet_LoadCompletedEventHandler();
	public delegate void Spreadsheet_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Spreadsheet_MouseOutEventHandler(Int32 button, Int32 shift, NetOffice.OWC10Api._Range target);
	public delegate void Spreadsheet_MouseOverEventHandler(Int32 button, Int32 shift, NetOffice.OWC10Api._Range target);
	public delegate void Spreadsheet_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Spreadsheet_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void Spreadsheet_SelectionChangeEventHandler();
	public delegate void Spreadsheet_SelectionChangingEventHandler(NetOffice.OWC10Api._Range range);
	public delegate void Spreadsheet_SheetActivateEventHandler(NetOffice.OWC10Api.Worksheet sh);
	public delegate void Spreadsheet_SheetCalculateEventHandler(NetOffice.OWC10Api.Worksheet sh);
	public delegate void Spreadsheet_SheetChangeEventHandler(NetOffice.OWC10Api.Worksheet sh, NetOffice.OWC10Api._Range target);
	public delegate void Spreadsheet_SheetDeactivateEventHandler(NetOffice.OWC10Api.Worksheet sh);
	public delegate void Spreadsheet_SheetFollowHyperlinkEventHandler(NetOffice.OWC10Api.Worksheet sh, NetOffice.OWC10Api.Hyperlink target);
	public delegate void Spreadsheet_StartEditEventHandler(ICOMObject selection, NetOffice.OWC10Api.ByRef initialValue, NetOffice.OWC10Api.ByRef cancel, NetOffice.OWC10Api.ByRef errorDescription);
	public delegate void Spreadsheet_ViewChangeEventHandler(NetOffice.OWC10Api._Range target);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Spreadsheet 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ISpreadsheetEventSink))]
	[TypeId("0002E551-0000-0000-C000-000000000046")]
    public interface Spreadsheet : ISpreadsheet, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_BeforeContextMenuEventHandler BeforeContextMenuEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_BeforeKeyDownEventHandler BeforeKeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_BeforeKeyPressEventHandler BeforeKeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_BeforeKeyUpEventHandler BeforeKeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_CommandEnabledEventHandler CommandEnabledEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_CommandCheckedEventHandler CommandCheckedEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_CommandTipTextEventHandler CommandTipTextEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_CommandExecuteEventHandler CommandExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_EndEditEventHandler EndEditEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_InitializeEventHandler InitializeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_KeyDownEventHandler KeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_KeyPressEventHandler KeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_KeyUpEventHandler KeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_LoadCompletedEventHandler LoadCompletedEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_MouseOutEventHandler MouseOutEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_MouseOverEventHandler MouseOverEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_MouseWheelEventHandler MouseWheelEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SelectionChangeEventHandler SelectionChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SelectionChangingEventHandler SelectionChangingEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SheetActivateEventHandler SheetActivateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SheetCalculateEventHandler SheetCalculateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SheetChangeEventHandler SheetChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SheetDeactivateEventHandler SheetDeactivateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_StartEditEventHandler StartEditEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event Spreadsheet_ViewChangeEventHandler ViewChangeEvent;

        #endregion
    }
}
