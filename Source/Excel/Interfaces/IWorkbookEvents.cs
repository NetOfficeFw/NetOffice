using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IWorkbookEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00024412-0001-0000-C000-000000000046")]
	public interface IWorkbookEvents : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Open();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Activate();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Deactivate();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 BeforeClose(bool cancel);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 BeforeSave(bool saveAsUI, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 BeforePrint(bool cancel);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 NewSheet(object sh);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 AddinInstall();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 AddinUninstall();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 WindowResize(NetOffice.ExcelApi.Window wn);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 WindowActivate(NetOffice.ExcelApi.Window wn);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 WindowDeactivate(NetOffice.ExcelApi.Window wn);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetSelectionChange(object sh, NetOffice.ExcelApi.Range target);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetBeforeDoubleClick(object sh, NetOffice.ExcelApi.Range target, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetBeforeRightClick(object sh, NetOffice.ExcelApi.Range target, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetActivate(object sh);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetDeactivate(object sh);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetCalculate(object sh);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetChange(object sh, NetOffice.ExcelApi.Range target);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Hyperlink target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 SheetFollowHyperlink(object sh, NetOffice.ExcelApi.Hyperlink target);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 SheetPivotTableUpdate(object sh, NetOffice.ExcelApi.PivotTable target);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 PivotTableCloseConnection(NetOffice.ExcelApi.PivotTable target);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 PivotTableOpenConnection(NetOffice.ExcelApi.PivotTable target);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 Sync(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="isRefresh">bool isRefresh</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 BeforeXmlImport(NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="isRefresh">bool isRefresh</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlImportResult result</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 AfterXmlImport(NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 BeforeXmlExport(NetOffice.ExcelApi.XmlMap map, string url, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlExportResult result</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 AfterXmlExport(NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="description">string description</param>
		/// <param name="sheet">string sheet</param>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 RowsetComplete(string description, string sheet, bool success);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="targetRange">NetOffice.ExcelApi.Range targetRange</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 SheetPivotTableAfterValueChange(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 SheetPivotTableBeforeAllocateChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 SheetPivotTableBeforeCommitChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 SheetPivotTableBeforeDiscardChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 SheetPivotTableChangeSync(object sh, NetOffice.ExcelApi.PivotTable target);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 AfterSave(bool success);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="ch">NetOffice.ExcelApi.Chart ch</param>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 NewChart(NetOffice.ExcelApi.Chart ch);

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
		/// <param name="changes">NetOffice.ExcelApi.ModelChanges changes</param>
		[SupportByVersion("Excel", 15, 16)]
		Int32 ModelChange(NetOffice.ExcelApi.ModelChanges changes);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 15, 16)]
		Int32 SheetBeforeDelete(object sh);

		#endregion
	}
}
