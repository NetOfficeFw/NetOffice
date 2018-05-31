using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Workbook_OpenEventHandler();
	public delegate void Workbook_ActivateEventHandler();
	public delegate void Workbook_DeactivateEventHandler();
	public delegate void Workbook_BeforeCloseEventHandler(ref bool cancel);
	public delegate void Workbook_BeforeSaveEventHandler(bool saveAsUI, ref bool cancel);
	public delegate void Workbook_BeforePrintEventHandler(ref bool cancel);
	public delegate void Workbook_NewSheetEventHandler(ICOMObject sh);
	public delegate void Workbook_AddinInstallEventHandler();
	public delegate void Workbook_AddinUninstallEventHandler();
	public delegate void Workbook_WindowResizeEventHandler(NetOffice.ExcelApi.Window wn);
	public delegate void Workbook_WindowActivateEventHandler(NetOffice.ExcelApi.Window wn);
	public delegate void Workbook_WindowDeactivateEventHandler(NetOffice.ExcelApi.Window wn);
	public delegate void Workbook_SheetSelectionChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
	public delegate void Workbook_SheetBeforeDoubleClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Workbook_SheetBeforeRightClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Workbook_SheetActivateEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetDeactivateEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetCalculateEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
	public delegate void Workbook_SheetFollowHyperlinkEventHandler(ICOMObject sh, NetOffice.ExcelApi.Hyperlink target);
	public delegate void Workbook_SheetPivotTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_PivotTableCloseConnectionEventHandler(NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_PivotTableOpenConnectionEventHandler(NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_SyncEventHandler(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Workbook_BeforeXmlImportEventHandler(NetOffice.ExcelApi.XmlMap map, string url, bool IsRefresh, ref bool cancel);
	public delegate void Workbook_AfterXmlImportEventHandler(NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result);
	public delegate void Workbook_BeforeXmlExportEventHandler(NetOffice.ExcelApi.XmlMap map, string url, ref bool cancel);
	public delegate void Workbook_AfterXmlExportEventHandler(NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result);
	public delegate void Workbook_RowsetCompleteEventHandler(string description, string sheet, bool success);
	public delegate void Workbook_SheetPivotTableAfterValueChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);
	public delegate void Workbook_SheetPivotTableBeforeAllocateChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Workbook_SheetPivotTableBeforeCommitChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Workbook_SheetPivotTableBeforeDiscardChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void Workbook_SheetPivotTableChangeSyncEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_AfterSaveEventHandler(bool success);
	public delegate void Workbook_NewChartEventHandler(NetOffice.ExcelApi.Chart ch);
	public delegate void Workbook_SheetLensGalleryRenderCompleteEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.TableObject target);
	public delegate void Workbook_ModelChangeEventHandler(NetOffice.ExcelApi.ModelChanges changes);
	public delegate void Workbook_SheetBeforeDeleteEventHandler(ICOMObject sh);
#pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Workbook
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835568.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventInterface(typeof(EventContracts.WorkbookEvents))]
    public interface Workbook : _Workbook, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196215.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_OpenEventHandler OpenEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823078.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_ActivateEventHandler ActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822521.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_DeactivateEventHandler DeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194765.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_BeforeCloseEventHandler BeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840057.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_BeforeSaveEventHandler BeforeSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195836.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_BeforePrintEventHandler BeforePrintEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821246.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_NewSheetEventHandler NewSheetEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822158.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_AddinInstallEventHandler AddinInstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840207.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_AddinUninstallEventHandler AddinUninstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822648.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_WindowResizeEventHandler WindowResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840441.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_WindowActivateEventHandler WindowActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839676.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_WindowDeactivateEventHandler WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837368.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetSelectionChangeEventHandler SheetSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822360.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839675.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195710.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetActivateEventHandler SheetActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838589.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetDeactivateEventHandler SheetDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193282.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
         event Workbook_SheetCalculateEventHandler SheetCalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196611.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetChangeEventHandler SheetChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838573.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion Excel 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193598.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Workbook_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231875.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Workbook_PivotTableCloseConnectionEventHandler PivotTableCloseConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231367.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Workbook_PivotTableOpenConnectionEventHandler PivotTableOpenConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231646.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Workbook_SyncEventHandler SyncEvent;

        /// <summary>
        /// SupportByVersion Excel 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837099.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Workbook_BeforeXmlImportEventHandler BeforeXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838071.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Workbook_AfterXmlImportEventHandler AfterXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840659.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Workbook_BeforeXmlExportEventHandler BeforeXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841244.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Workbook_AfterXmlExportEventHandler AfterXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193275.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        event Workbook_RowsetCompleteEventHandler RowsetCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834987.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196066.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834653.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840414.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838762.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_SheetPivotTableChangeSyncEventHandler SheetPivotTableChangeSyncEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836466.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_AfterSaveEventHandler AfterSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823186.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Workbook_NewChartEventHandler NewChartEvent;

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230636.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Workbook_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229546.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Workbook_SheetTableUpdateEventHandler SheetTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232031.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Workbook_ModelChangeEventHandler ModelChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448396.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Workbook_SheetBeforeDeleteEventHandler SheetBeforeDeleteEvent;

        #endregion
    }
}
