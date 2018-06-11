using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Chart_ActivateEventHandler();
	public delegate void Chart_DeactivateEventHandler();
	public delegate void Chart_ResizeEventHandler();
	public delegate void Chart_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Chart_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Chart_MouseMoveEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Chart_BeforeRightClickEventHandler(ref bool cancel);
	public delegate void Chart_DragPlotEventHandler();
	public delegate void Chart_DragOverEventHandler();
	public delegate void Chart_BeforeDoubleClickEventHandler(Int32 elementID, Int32 arg1, Int32 arg2, ref bool cancel);
	public delegate void Chart_SelectEventHandler(Int32 elementID, Int32 arg1, Int32 arg2);
	public delegate void Chart_SeriesChangeEventHandler(Int32 seriesIndex, Int32 pointIndex);
	public delegate void Chart_CalculateEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Chart
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194426.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ChartEvents))]
	[TypeId("00020821-0000-0000-C000-000000000046")]
    public interface Chart : _Chart, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834456.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_ActivateEventHandler ActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838241.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_DeactivateEventHandler DeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839406.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_ResizeEventHandler ResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822567.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_MouseDownEventHandler MouseDownEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197532.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_MouseUpEventHandler MouseUpEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837995.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_MouseMoveEventHandler MouseMoveEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839270.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_BeforeRightClickEventHandler BeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_DragPlotEventHandler DragPlotEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_DragOverEventHandler DragOverEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197223.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_BeforeDoubleClickEventHandler BeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192964.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_SelectEventHandler SelectEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834746.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_SeriesChangeEventHandler SeriesChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820890.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Chart_CalculateEventHandler CalculateEvent;

        #endregion
    }
}
