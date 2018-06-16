using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	#region Delegates

	#pragma warning disable
	public delegate void ChartSpace_DataSetChangeEventHandler();
	public delegate void ChartSpace_DblClickEventHandler();
	public delegate void ChartSpace_ClickEventHandler();
	public delegate void ChartSpace_KeyDownEventHandler(Int32 keyCode, Int32 shift);
	public delegate void ChartSpace_KeyUpEventHandler(Int32 keyCode, Int32 shift);
	public delegate void ChartSpace_KeyPressEventHandler(Int32 keyAscii);
	public delegate void ChartSpace_BeforeKeyDownEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void ChartSpace_BeforeKeyUpEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void ChartSpace_BeforeKeyPressEventHandler(Int32 keyAscii, NetOffice.OWC10Api.ByRef Cancel);
	public delegate void ChartSpace_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void ChartSpace_MouseMoveEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void ChartSpace_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void ChartSpace_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void ChartSpace_SelectionChangeEventHandler();
	public delegate void ChartSpace_BeforeScreenTipEventHandler(NetOffice.OWC10Api.ByRef tipText, COMObject contextObject);
	public delegate void ChartSpace_CommandEnabledEventHandler(object command, NetOffice.OWC10Api.ByRef enabled);
	public delegate void ChartSpace_CommandCheckedEventHandler(object command, NetOffice.OWC10Api.ByRef Checked);
	public delegate void ChartSpace_CommandTipTextEventHandler(object command, NetOffice.OWC10Api.ByRef caption);
	public delegate void ChartSpace_CommandBeforeExecuteEventHandler(object command, NetOffice.OWC10Api.ByRef Cancel);
	public delegate void ChartSpace_CommandExecuteEventHandler(object command, bool succeeded);
	public delegate void ChartSpace_BeforeContextMenuEventHandler(Int32 x, Int32 y, NetOffice.OWC10Api.ByRef menu, NetOffice.OWC10Api.ByRef cancel);
	public delegate void ChartSpace_BeforeRenderEventHandler(NetOffice.OWC10Api.ChChartDraw drawObject, COMObject chartObject, NetOffice.OWC10Api.ByRef cancel);
	public delegate void ChartSpace_AfterRenderEventHandler(NetOffice.OWC10Api.ChChartDraw drawObject, COMObject chartObject);
	public delegate void ChartSpace_AfterFinalRenderEventHandler(NetOffice.OWC10Api.ChChartDraw drawObject);
	public delegate void ChartSpace_AfterLayoutEventHandler(NetOffice.OWC10Api.ChChartDraw drawObject);
	public delegate void ChartSpace_ViewChangeEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass ChartSpace 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.IChartEvents))]
	[TypeId("0002E556-0000-0000-C000-000000000046")]
    public interface ChartSpace : ChChartSpace, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_DataSetChangeEventHandler DataSetChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_DblClickEventHandler DblClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_ClickEventHandler ClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_KeyDownEventHandler KeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_KeyUpEventHandler KeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_KeyPressEventHandler KeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_BeforeKeyDownEventHandler BeforeKeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_BeforeKeyUpEventHandler BeforeKeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_BeforeKeyPressEventHandler BeforeKeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_MouseWheelEventHandler MouseWheelEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_SelectionChangeEventHandler SelectionChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_BeforeScreenTipEventHandler BeforeScreenTipEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_CommandEnabledEventHandler CommandEnabledEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_CommandCheckedEventHandler CommandCheckedEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_CommandTipTextEventHandler CommandTipTextEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_CommandExecuteEventHandler CommandExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_BeforeContextMenuEventHandler BeforeContextMenuEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_BeforeRenderEventHandler BeforeRenderEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_AfterRenderEventHandler AfterRenderEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_AfterFinalRenderEventHandler AfterFinalRenderEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_AfterLayoutEventHandler AfterLayoutEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event ChartSpace_ViewChangeEventHandler ViewChangeEvent;

        #endregion
    }
}
