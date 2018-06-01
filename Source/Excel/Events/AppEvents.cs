using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.EventContracts
{
    /// <summary>
    /// AppEvents
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface AppEvents
    {
        /// <summary>
        /// NewWorkbook
        /// </summary>
        /// <param name="wb"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1565)]
        void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        /// <summary>
        /// SheetSelectionChange
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
        void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// SheetBeforeDoubleClick
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
        void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        /// <summary>
        /// SheetBeforeRightClick
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
        void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        /// <summary>
        /// SheetActivate
        /// </summary>
        /// <param name="sh"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
        void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetDeactivate
        /// </summary>
        /// <param name="sh"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
        void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetCalculate
        /// </summary>
        /// <param name="sh"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
        void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetChange
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
        void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// WorkbookOpen
        /// </summary>
        /// <param name="wb"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1567)]
        void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        /// <summary>
        /// WorkbookActivate
        /// </summary>
        /// <param name="wb"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1568)]
        void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        /// <summary>
        /// WorkbookDeactivate
        /// </summary>
        /// <param name="wb"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1569)]
        void WorkbookDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        /// <summary>
        /// WorkbookBeforeClose
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1570)]
        void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel);

        /// <summary>
        /// WorkbookBeforeSave
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1571)]
        void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object saveAsUI, [In] [Out] ref object cancel);

        /// <summary>
        /// WorkbookBeforePrint
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1572)]
        void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel);

        /// <summary>
        /// WorkbookNewSheet
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sh"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1573)]
        void WorkbookNewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// WorkbookAddinInstall
        /// </summary>
        /// <param name="wb"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1574)]
        void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        /// <summary>
        /// WorkbookAddinUninstall
        /// </summary>
        /// <param name="wb"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1575)]
        void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        /// <summary>
        /// WindowResize
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="wn"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
        void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowActivate
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="wn"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
        void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowDeactivate
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="wn"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
        void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// SheetFollowHyperlink
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Hyperlink))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
        void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// SheetPivotTableUpdate
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
        void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// WorkbookPivotTableCloseConnection
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2160)]
        void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// WorkbookPivotTableOpenConnection
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2161)]
        void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// WorkbookSync
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="syncEventType"></param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2289)]
        void WorkbookSync([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object syncEventType);

        /// <summary>
        /// WorkbookBeforeXmlImport
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="url"></param>
        /// <param name="isRefresh"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("isRefresh", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2290)]
        void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel);

        /// <summary>
        /// WorkbookAfterXmlImport
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="isRefresh"></param>
        /// <param name="result"></param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("isRefresh", SinkArgumentType.Bool)]
        [SinkArgument("result", SinkArgumentType.Enum, typeof(ExcelApi.Enums.XlXmlExportResult))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2291)]
        void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result);

        /// <summary>
        /// WorkbookBeforeXmlExport
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="url"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2292)]
        void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel);

        /// <summary>
        /// WorkbookAfterXmlExport
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="url"></param>
        /// <param name="result"></param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("result", SinkArgumentType.Enum, typeof(NetOffice.ExcelApi.Enums.XlXmlExportResult))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2293)]
        void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result);

        /// <summary>
        /// WorkbookRowsetComplete
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="description"></param>
        /// <param name="sheet"></param>
        /// <param name="success"></param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("sheet", SinkArgumentType.String)]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2611)]
        void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object description, [In] object sheet, [In] object success);

        /// <summary>
        /// AfterCalculate
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2612)]
        void AfterCalculate();

        /// <summary>
        /// SheetPivotTableAfterValueChange
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="targetRange"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("targetRange", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
        void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

        /// <summary>
        /// SheetPivotTableBeforeAllocateChanges
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="valueChangeStart"></param>
        /// <param name="valueChangeEnd"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
        void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

        /// <summary>
        /// SheetPivotTableBeforeCommitChanges
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="valueChangeStart"></param>
        /// <param name="valueChangeEnd"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
        void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

        /// <summary>
        /// SheetPivotTableBeforeDiscardChanges
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="valueChangeStart"></param>
        /// <param name="valueChangeEnd"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
        void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

        /// <summary>
        /// ProtectedViewWindowOpen
        /// </summary>
        /// <param name="pvw"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2903)]
        void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        /// <summary>
        /// ProtectedViewWindowBeforeEdit
        /// </summary>
        /// <param name="pvw"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2905)]
        void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowBeforeClose
        /// </summary>
        /// <param name="pvw"></param>
        /// <param name="reason"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [SinkArgument("reason", SinkArgumentType.Enum, typeof(ExcelApi.Enums.XlProtectedViewCloseReason))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2906)]
        void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] object reason, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowResize
        /// </summary>
        /// <param name="pvw"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2908)]
        void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        /// <summary>
        /// ProtectedViewWindowActivate
        /// </summary>
        /// <param name="pvw"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2909)]
        void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        /// <summary>
        /// ProtectedViewWindowDeactivate
        /// </summary>
        /// <param name="pvw"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2910)]
        void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        /// <summary>
        /// WorkbookAfterSave
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="success"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2911)]
        void WorkbookAfterSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object success);

        /// <summary>
        /// WorkbookNewChart
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="ch"></param>
        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("ch", typeof(ExcelApi.Chart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2912)]
        void WorkbookNewChart([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object ch);

        /// <summary>
        /// SheetLensGalleryRenderComplete
        /// </summary>
        /// <param name="sh"></param>
        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3073)]
        void SheetLensGalleryRenderComplete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        /// <summary>
        /// SheetTableUpdate
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.TableObject))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3074)]
        void SheetTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        /// <summary>
        /// WorkbookModelChange
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="changes"></param>
        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("changes", typeof(ExcelApi.ModelChanges))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3078)]
        void WorkbookModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object changes);

        /// <summary>
        /// SheetBeforeDelete
        /// </summary>
        /// <param name="sh"></param>
        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3077)]
        void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
    }
}
