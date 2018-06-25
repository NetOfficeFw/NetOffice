using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
    #region Delegates

    #pragma warning disable
    public delegate void Application_NewWorkbookEventHandler(NetOffice.ExcelApi.Workbook wb);
    public delegate void Application_SheetSelectionChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
    public delegate void Application_SheetBeforeDoubleClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
    public delegate void Application_SheetBeforeRightClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
    public delegate void Application_SheetActivateEventHandler(ICOMObject sh);
    public delegate void Application_SheetDeactivateEventHandler(ICOMObject sh);
    public delegate void Application_SheetCalculateEventHandler(ICOMObject sh);
    public delegate void Application_SheetChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
    public delegate void Application_WorkbookOpenEventHandler(NetOffice.ExcelApi.Workbook wb);
    public delegate void Application_WorkbookActivateEventHandler(NetOffice.ExcelApi.Workbook wb);
    public delegate void Application_WorkbookDeactivateEventHandler(NetOffice.ExcelApi.Workbook wb);
    public delegate void Application_WorkbookBeforeCloseEventHandler(NetOffice.ExcelApi.Workbook wb, ref bool cancel);
    public delegate void Application_WorkbookBeforeSaveEventHandler(NetOffice.ExcelApi.Workbook wb, bool saveAsUI, ref bool cancel);
    public delegate void Application_WorkbookBeforePrintEventHandler(NetOffice.ExcelApi.Workbook wb, ref bool cancel);
    public delegate void Application_WorkbookNewSheetEventHandler(NetOffice.ExcelApi.Workbook wb, ICOMObject sh);
    public delegate void Application_WorkbookAddinInstallEventHandler(NetOffice.ExcelApi.Workbook wb);
    public delegate void Application_WorkbookAddinUninstallEventHandler(NetOffice.ExcelApi.Workbook wb);
    public delegate void Application_WindowResizeEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);
    public delegate void Application_WindowActivateEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);
    public delegate void Application_WindowDeactivateEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);
    public delegate void Application_SheetFollowHyperlinkEventHandler(ICOMObject sh, NetOffice.ExcelApi.Hyperlink target);
    public delegate void Application_SheetPivotTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable target);
    public delegate void Application_WorkbookPivotTableCloseConnectionEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target);
    public delegate void Application_WorkbookPivotTableOpenConnectionEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target);
    public delegate void Application_WorkbookSyncEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
    public delegate void Application_WorkbookBeforeXmlImportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, ref bool cancel);
    public delegate void Application_WorkbookAfterXmlImportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result);
    public delegate void Application_WorkbookBeforeXmlExportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, ref bool cancel);
    public delegate void Application_WorkbookAfterXmlExportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result);
    public delegate void Application_WorkbookRowsetCompleteEventHandler(NetOffice.ExcelApi.Workbook wb, string description, string sheet, bool success);
    public delegate void Application_AfterCalculateEventHandler();
    public delegate void Application_SheetPivotTableAfterValueChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);
    public delegate void Application_SheetPivotTableBeforeAllocateChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
    public delegate void Application_SheetPivotTableBeforeCommitChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
    public delegate void Application_SheetPivotTableBeforeDiscardChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
    public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
    public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw, ref bool cancel);
    public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason, ref bool cancel);
    public delegate void Application_ProtectedViewWindowResizeEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
    public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
    public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
    public delegate void Application_WorkbookAfterSaveEventHandler(NetOffice.ExcelApi.Workbook wb, bool success);
    public delegate void Application_WorkbookNewChartEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Chart ch);
    public delegate void Application_SheetLensGalleryRenderCompleteEventHandler(ICOMObject sh);
    public delegate void Application_SheetTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.TableObject target);
    public delegate void Application_WorkbookModelChangeEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.ModelChanges changes);
    public delegate void Application_SheetBeforeDeleteEventHandler(ICOMObject sh);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.ExcelApi.Behind.Application
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.ExcelApi.Behind.Application
    {
        private string _defaultProgId = "Excel.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Excel
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Excel based on given id.
        /// This can be used to target a specific version of Microsoft Excel.
        /// Example usage:
        /// "Microsoft.Excel.12" to target Excel 2007
        /// "Microsoft.Excel.14" to target Excel 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Excel
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Excel
        /// </summary>
        /// <param name="mode">indicates where is the call coming from</param>
        public ApplicationClass(NetOffice.Callers.InteropCompatibilityClassCreateMode mode)
        {
            if (mode == NetOffice.Callers.InteropCompatibilityClassCreateMode.Direct)
            {
                ICOMObjectInitialize init = (ICOMObjectInitialize)this;
                init.InitializeCOMObject(_defaultProgId);
            }
        }
    }

    /// <summary>
    /// CoClass Application
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass), ComProgId("Excel.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.ExcelApi.EventContracts.AppEvents))]
	[TypeId("00024500-0000-0000-C000-000000000046")]
    public interface Application : _Application, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
    {
        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837373.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_NewWorkbookEventHandler NewWorkbookEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839035.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetSelectionChangeEventHandler SheetSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836225.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840532.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193288.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetActivateEventHandler SheetActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823120.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetDeactivateEventHandler SheetDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835607.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetCalculateEventHandler SheetCalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193591.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetChangeEventHandler SheetChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196583.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookOpenEventHandler WorkbookOpenEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837347.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookActivateEventHandler WorkbookActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193560.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookDeactivateEventHandler WorkbookDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836770.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookBeforeCloseEventHandler WorkbookBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840422.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookBeforeSaveEventHandler WorkbookBeforeSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195507.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookBeforePrintEventHandler WorkbookBeforePrintEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198367.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookNewSheetEventHandler WorkbookNewSheetEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836206.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookAddinInstallEventHandler WorkbookAddinInstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835570.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookAddinUninstallEventHandler WorkbookAddinUninstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836166.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowResizeEventHandler WindowResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821328.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowActivateEventHandler WindowActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822473.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowDeactivateEventHandler WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821956.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840950.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Application_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198029.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821547.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        event Application_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839042.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Application_WorkbookSyncEventHandler WorkbookSyncEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196324.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Application_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837416.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Application_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195824.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Application_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836803.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        event Application_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839165.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        event Application_WorkbookRowsetCompleteEventHandler WorkbookRowsetCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840621.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        event Application_AfterCalculateEventHandler AfterCalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193316.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838226.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838379.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835217.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194431.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838239.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821579.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836848.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195451.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196820.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198184.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_WorkbookAfterSaveEventHandler WorkbookAfterSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834985.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        event Application_WorkbookNewChartEventHandler WorkbookNewChartEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227506.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Application_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229805.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Application_SheetTableUpdateEventHandler SheetTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229611.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Application_WorkbookModelChangeEventHandler WorkbookModelChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448391.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        event Application_SheetBeforeDeleteEventHandler SheetBeforeDeleteEvent;

        #endregion
    }
}
