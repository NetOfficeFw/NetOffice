using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Worksheet_SelectionChangeEventHandler(NetOffice.ExcelApi.Range target);
	public delegate void Worksheet_BeforeDoubleClickEventHandler(NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Worksheet_BeforeRightClickEventHandler(NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Worksheet_ActivateEventHandler();
	public delegate void Worksheet_DeactivateEventHandler();
	public delegate void Worksheet_CalculateEventHandler();
	public delegate void Worksheet_ChangeEventHandler(NetOffice.ExcelApi.Range target);
	public delegate void Worksheet_FollowHyperlinkEventHandler(NetOffice.ExcelApi.Hyperlink target);
	public delegate void Worksheet_PivotTableUpdateEventHandler(NetOffice.ExcelApi.PivotTable target);
	public delegate void Worksheet_PivotTableAfterValueChangeEventHandler(NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);
	public delegate void Worksheet_PivotTableBeforeAllocateChangesEventHandler(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Worksheet_PivotTableBeforeCommitChangesEventHandler(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Worksheet_PivotTableBeforeDiscardChangesEventHandler(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void Worksheet_PivotTableChangeSyncEventHandler(NetOffice.ExcelApi.PivotTable target);
	public delegate void Worksheet_LensGalleryRenderCompleteEventHandler();
	public delegate void Worksheet_TableUpdateEventHandler(NetOffice.ExcelApi.TableObject target);
	public delegate void Worksheet_BeforeDeleteEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Worksheet
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194464.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DocEvents))]
	[TypeId("00020820-0000-0000-C000-000000000046")]
    public interface Worksheet : _Worksheet, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194470.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_SelectionChangeEventHandler SelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196564.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_BeforeDoubleClickEventHandler BeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192993.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_BeforeRightClickEventHandler BeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198220.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_ActivateEventHandler ActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197183.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_DeactivateEventHandler DeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838823.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_CalculateEventHandler CalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839775.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_ChangeEventHandler ChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838843.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Worksheet_FollowHyperlinkEventHandler FollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822105.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Worksheet_PivotTableUpdateEventHandler PivotTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193517.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Worksheet_PivotTableAfterValueChangeEventHandler PivotTableAfterValueChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195070.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Worksheet_PivotTableBeforeAllocateChangesEventHandler PivotTableBeforeAllocateChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198138.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Worksheet_PivotTableBeforeCommitChangesEventHandler PivotTableBeforeCommitChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836187.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Worksheet_PivotTableBeforeDiscardChangesEventHandler PivotTableBeforeDiscardChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838251.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Worksheet_PivotTableChangeSyncEventHandler PivotTableChangeSyncEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227603.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Worksheet_LensGalleryRenderCompleteEventHandler LensGalleryRenderCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229788.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Worksheet_TableUpdateEventHandler TableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448393.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Worksheet_BeforeDeleteEventHandler BeforeDeleteEvent;

        #endregion
    }
}
