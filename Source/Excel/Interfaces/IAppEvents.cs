using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IAppEvents 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IAppEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 NewWorkbook(NetOffice.ExcelApi.Workbook wb);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetSelectionChange(object sh, NetOffice.ExcelApi.Range target);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetBeforeDoubleClick(object sh, NetOffice.ExcelApi.Range target, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetBeforeRightClick(object sh, NetOffice.ExcelApi.Range target, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetActivate(object sh);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetDeactivate(object sh);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetCalculate(object sh);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetChange(object sh, NetOffice.ExcelApi.Range target);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookOpen(NetOffice.ExcelApi.Workbook wb);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookActivate(NetOffice.ExcelApi.Workbook wb);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookDeactivate(NetOffice.ExcelApi.Workbook wb);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookBeforeClose(NetOffice.ExcelApi.Workbook wb, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="saveAsUI">bool saveAsUI</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookBeforeSave(NetOffice.ExcelApi.Workbook wb, bool saveAsUI, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookBeforePrint(NetOffice.ExcelApi.Workbook wb, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="sh">object sh</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookNewSheet(NetOffice.ExcelApi.Workbook wb, object sh);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookAddinInstall(NetOffice.ExcelApi.Workbook wb);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookAddinUninstall(NetOffice.ExcelApi.Workbook wb);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="wn">NetOffice.ExcelApi.Window wn</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WindowResize(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="wn">NetOffice.ExcelApi.Window wn</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WindowActivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="wn">NetOffice.ExcelApi.Window wn</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 WindowDeactivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.Hyperlink target</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SheetFollowHyperlink(object sh, NetOffice.ExcelApi.Hyperlink target);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 SheetPivotTableUpdate(object sh, NetOffice.ExcelApi.PivotTable target);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookPivotTableCloseConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 WorkbookPivotTableOpenConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 WorkbookSync(NetOffice.ExcelApi.Workbook wb, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
        /// <param name="url">string url</param>
        /// <param name="isRefresh">bool isRefresh</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 WorkbookBeforeXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
        /// <param name="isRefresh">bool isRefresh</param>
        /// <param name="result">NetOffice.ExcelApi.Enums.XlXmlImportResult result</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 WorkbookAfterXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
        /// <param name="url">string url</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 WorkbookBeforeXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
        /// <param name="url">string url</param>
        /// <param name="result">NetOffice.ExcelApi.Enums.XlXmlExportResult result</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 WorkbookAfterXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="description">string description</param>
        /// <param name="sheet">string sheet</param>
        /// <param name="success">bool success</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 WorkbookRowsetComplete(NetOffice.ExcelApi.Workbook wb, string description, string sheet, bool success);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 AfterCalculate();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="targetRange">NetOffice.ExcelApi.Range targetRange</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 SheetPivotTableAfterValueChange(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="valueChangeStart">Int32 valueChangeStart</param>
        /// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 SheetPivotTableBeforeAllocateChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="valueChangeStart">Int32 valueChangeStart</param>
        /// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 SheetPivotTableBeforeCommitChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="valueChangeStart">Int32 valueChangeStart</param>
        /// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 SheetPivotTableBeforeDiscardChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ProtectedViewWindowOpen(NetOffice.ExcelApi.ProtectedViewWindow pvw);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ProtectedViewWindowBeforeEdit(NetOffice.ExcelApi.ProtectedViewWindow pvw, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
        /// <param name="reason">NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ProtectedViewWindowBeforeClose(NetOffice.ExcelApi.ProtectedViewWindow pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ProtectedViewWindowResize(NetOffice.ExcelApi.ProtectedViewWindow pvw);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ProtectedViewWindowActivate(NetOffice.ExcelApi.ProtectedViewWindow pvw);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ProtectedViewWindowDeactivate(NetOffice.ExcelApi.ProtectedViewWindow pvw);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="success">bool success</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 WorkbookAfterSave(NetOffice.ExcelApi.Workbook wb, bool success);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="ch">NetOffice.ExcelApi.Chart ch</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 WorkbookNewChart(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Chart ch);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="sh">object sh</param>
        [SupportByVersion("Excel", 15, 16)]
        Int32 SheetLensGalleryRenderComplete(object sh);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="sh">object sh</param>
        /// <param name="target">NetOffice.ExcelApi.TableObject target</param>
        [SupportByVersion("Excel", 15, 16)]
        Int32 SheetTableUpdate(object sh, NetOffice.ExcelApi.TableObject target);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
        /// <param name="changes">NetOffice.ExcelApi.ModelChanges changes</param>
        [SupportByVersion("Excel", 15, 16)]
        Int32 WorkbookModelChange(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.ModelChanges changes);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="sh">object sh</param>
        [SupportByVersion("Excel", 15, 16)]
        Int32 SheetBeforeDelete(object sh);

        #endregion
    }
}
