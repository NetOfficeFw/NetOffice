using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	#region Delegates

	#pragma warning disable
	public delegate void PivotTable_SelectionChangeEventHandler();
	public delegate void PivotTable_ViewChangeEventHandler(NetOffice.OWC10Api.Enums.PivotViewReasonEnum reason);
	public delegate void PivotTable_DataChangeEventHandler(NetOffice.OWC10Api.Enums.PivotDataReasonEnum reason);
	public delegate void PivotTable_PivotTableChangeEventHandler(NetOffice.OWC10Api.Enums.PivotTableReasonEnum reason);
	public delegate void PivotTable_BeforeQueryEventHandler();
	public delegate void PivotTable_QueryEventHandler();
	public delegate void PivotTable_OnConnectEventHandler();
	public delegate void PivotTable_OnDisconnectEventHandler();
	public delegate void PivotTable_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void PivotTable_MouseMoveEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void PivotTable_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void PivotTable_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void PivotTable_ClickEventHandler();
	public delegate void PivotTable_DblClickEventHandler();
	public delegate void PivotTable_CommandEnabledEventHandler(object command, NetOffice.OWC10Api.ByRef enabled);
	public delegate void PivotTable_CommandCheckedEventHandler(object command, NetOffice.OWC10Api.ByRef Checked);
	public delegate void PivotTable_CommandTipTextEventHandler(object command, NetOffice.OWC10Api.ByRef caption);
	public delegate void PivotTable_CommandBeforeExecuteEventHandler(object command, NetOffice.OWC10Api.ByRef Cancel);
	public delegate void PivotTable_CommandExecuteEventHandler(object command, bool succeeded);
	public delegate void PivotTable_KeyDownEventHandler(Int32 keyCode, Int32 shift);
	public delegate void PivotTable_KeyUpEventHandler(Int32 keyCode, Int32 shift);
	public delegate void PivotTable_KeyPressEventHandler(Int32 keyAscii);
	public delegate void PivotTable_BeforeKeyDownEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_BeforeKeyUpEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_BeforeKeyPressEventHandler(Int32 keyAscii, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_BeforeContextMenuEventHandler(Int32 x, Int32 y, NetOffice.OWC10Api.ByRef Menu, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_StartEditEventHandler(ICOMObject selection, ICOMObject activeObject, NetOffice.OWC10Api.ByRef initialValue, NetOffice.OWC10Api.ByRef arrowMode, NetOffice.OWC10Api.ByRef caretPosition, NetOffice.OWC10Api.ByRef cancel, NetOffice.OWC10Api.ByRef errorDescription);
	public delegate void PivotTable_EndEditEventHandler(bool accept, NetOffice.OWC10Api.ByRef finalValue, NetOffice.OWC10Api.ByRef cancel, NetOffice.OWC10Api.ByRef errorDescription);
	public delegate void PivotTable_BeforeScreenTipEventHandler(NetOffice.OWC10Api.ByRef screenTipText, ICOMObject sourceObject);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass PivotTable 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.IPivotControlEvents))]
	[TypeId("0002E552-0000-0000-C000-000000000046")]
    public interface PivotTable : IPivotControl, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_SelectionChangeEventHandler SelectionChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_ViewChangeEventHandler ViewChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_DataChangeEventHandler DataChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_PivotTableChangeEventHandler PivotTableChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_BeforeQueryEventHandler BeforeQueryEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_QueryEventHandler QueryEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_OnConnectEventHandler OnConnectEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_OnDisconnectEventHandler OnDisconnectEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_MouseWheelEventHandler MouseWheelEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_CommandEnabledEventHandler CommandEnabledEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_CommandCheckedEventHandler CommandCheckedEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_CommandTipTextEventHandler CommandTipTextEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_CommandExecuteEventHandler CommandExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_KeyDownEventHandler KeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_KeyUpEventHandler KeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_KeyPressEventHandler KeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_BeforeKeyDownEventHandler BeforeKeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_BeforeKeyUpEventHandler BeforeKeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_BeforeKeyPressEventHandler BeforeKeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_BeforeContextMenuEventHandler BeforeContextMenuEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_StartEditEventHandler StartEditEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_EndEditEventHandler EndEditEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event PivotTable_BeforeScreenTipEventHandler BeforeScreenTipEvent;

        #endregion
    }
}
