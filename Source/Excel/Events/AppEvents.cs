using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Events
{
    #pragma warning disable

    #region SinkPoint Interface

    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface AppEvents
    {
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1565)]
        void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
        void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
        void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
        void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
        void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
        void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
        void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
        void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1567)]
        void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1568)]
        void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1569)]
        void WorkbookDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1570)]
        void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1571)]
        void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object saveAsUI, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1572)]
        void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1573)]
        void WorkbookNewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1574)]
        void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1575)]
        void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
        void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
        void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
        void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Hyperlink))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
        void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
        void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2160)]
        void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2161)]
        void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2289)]
        void WorkbookSync([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object syncEventType);

        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("isRefresh", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2290)]
        void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("isRefresh", SinkArgumentType.Bool)]
        [SinkArgument("result", SinkArgumentType.Enum, typeof(ExcelApi.Enums.XlXmlExportResult))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2291)]
        void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result);

        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2292)]
        void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("result", SinkArgumentType.Enum, typeof(NetOffice.ExcelApi.Enums.XlXmlExportResult))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2293)]
        void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result);

        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("sheet", SinkArgumentType.String)]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2611)]
        void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object description, [In] object sheet, [In] object success);

        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2612)]
        void AfterCalculate();

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("targetRange", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
        void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
        void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
        void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
        void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2903)]
        void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2905)]
        void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [SinkArgument("reason", SinkArgumentType.Enum, typeof(ExcelApi.Enums.XlProtectedViewCloseReason))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2906)]
        void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] object reason, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2908)]
        void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2909)]
        void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("pvw", typeof(ExcelApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2910)]
        void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2911)]
        void WorkbookAfterSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object success);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("ch", typeof(ExcelApi.Chart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2912)]
        void WorkbookNewChart([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object ch);

        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3073)]
        void SheetLensGalleryRenderComplete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.TableObject))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3074)]
        void SheetTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("wb", typeof(ExcelApi.Workbook))]
        [SinkArgument("changes", typeof(ExcelApi.ModelChanges))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3078)]
        void WorkbookModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object changes);

        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3077)]
        void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
    }

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class AppEvents_SinkHelper : SinkHelper, AppEvents
    {
        #region Static

        public static readonly string Id = "00024413-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public AppEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {  
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region AppEvents

        public void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("NewWorkbook"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }
            
            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("NewWorkbook", ref paramsArray);
        }

        public void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }
            
            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetSelectionChange", ref paramsArray);
        }

        public void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetBeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(sh, target, cancel);
                return;
            }
            
            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("SheetBeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetBeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(sh, target, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("SheetBeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetActivate"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetActivate", ref paramsArray);
        }

        public void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetDeactivate"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }
           
            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetDeactivate", ref paramsArray);
        }

        public void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetCalculate"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetCalculate", ref paramsArray);
        }

        public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetChange"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetChange", ref paramsArray);
        }

        public void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookOpen"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookOpen", ref paramsArray);
        }

        public void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookActivate"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookActivate", ref paramsArray);
        }

        public void WorkbookDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookDeactivate"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookDeactivate", ref paramsArray);
        }

        public void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeClose"))
            {
                Invoker.ReleaseParamsArray(wb, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WorkbookBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object saveAsUI, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeSave"))
            {
                Invoker.ReleaseParamsArray(wb, saveAsUI, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            bool newSaveAsUI = ToBoolean(saveAsUI);
            object[] paramsArray = new object[3];
            paramsArray[0] = newWb;
            paramsArray[1] = newSaveAsUI;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("WorkbookBeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeSave"))
            {
                Invoker.ReleaseParamsArray(wb, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WorkbookBeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void WorkbookNewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("WorkbookNewSheet"))
            {
                Invoker.ReleaseParamsArray(wb, sh);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newSh;
            EventBinding.RaiseCustomEvent("WorkbookNewSheet", ref paramsArray);
        }

        public void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookAddinInstall"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookAddinInstall", ref paramsArray);
        }

        public void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookAddinUninstall"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookAddinUninstall", ref paramsArray);
        }

        public void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowResize"))
            {
                Invoker.ReleaseParamsArray(wb, wn);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wb, NetOffice.ExcelApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowResize", ref paramsArray);
        }

        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(wb, wn);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wb, NetOffice.ExcelApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
        }

        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(wb, wn);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wb, NetOffice.ExcelApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
        }

        public void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetFollowHyperlink"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Hyperlink newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Hyperlink>(EventClass, target, NetOffice.ExcelApi.Hyperlink.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetFollowHyperlink", ref paramsArray);
        }

        public void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetPivotTableUpdate"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetPivotTableUpdate", ref paramsArray);
        }

        public void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("WorkbookPivotTableCloseConnection"))
            {
                Invoker.ReleaseParamsArray(wb, target);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("WorkbookPivotTableCloseConnection", ref paramsArray);
        }

        public void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("WorkbookPivotTableOpenConnection"))
            {
                Invoker.ReleaseParamsArray(wb, target);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("WorkbookPivotTableOpenConnection", ref paramsArray);
        }

        public void WorkbookSync([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object syncEventType)
        {
            if (!Validate("WorkbookSync"))
            {
                Invoker.ReleaseParamsArray(wb, syncEventType);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newSyncEventType;
            EventBinding.RaiseCustomEvent("WorkbookSync", ref paramsArray);
        }

        public void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeXmlImport"))
            {
                Invoker.ReleaseParamsArray(wb, map, url, isRefresh, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, NetOffice.ExcelApi.XmlMap.LateBindingApiWrapperType);
            string newUrl = Convert.ToString(url);
            bool newIsRefresh = ToBoolean(isRefresh);
            object[] paramsArray = new object[5];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newUrl;
            paramsArray[3] = newIsRefresh;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("WorkbookBeforeXmlImport", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        public void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result)
        {
            if (!Validate("WorkbookAfterXmlImport"))
            {
                Invoker.ReleaseParamsArray(wb, map, isRefresh, result);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, NetOffice.ExcelApi.XmlMap.LateBindingApiWrapperType);
            bool newIsRefresh = ToBoolean(isRefresh);
            NetOffice.ExcelApi.Enums.XlXmlImportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlImportResult)result;
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newIsRefresh;
            paramsArray[3] = newResult;
            EventBinding.RaiseCustomEvent("WorkbookAfterXmlImport", ref paramsArray);
        }

        public void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeXmlExport"))
            {
                Invoker.ReleaseParamsArray(wb, map, url, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, NetOffice.ExcelApi.XmlMap.LateBindingApiWrapperType);
            string newUrl = ToString(url);
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newUrl;
            paramsArray.SetValue(cancel, 3);
            EventBinding.RaiseCustomEvent("WorkbookBeforeXmlExport", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
        }

        public void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result)
        {
            if (!Validate("WorkbookAfterXmlExport"))
            {
                Invoker.ReleaseParamsArray(wb, map, url, result);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, NetOffice.ExcelApi.XmlMap.LateBindingApiWrapperType);
            string newUrl = ToString(url);
            NetOffice.ExcelApi.Enums.XlXmlExportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlExportResult)result;
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newUrl;
            paramsArray[3] = newResult;
            EventBinding.RaiseCustomEvent("WorkbookAfterXmlExport", ref paramsArray);
        }
        
        public void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object description, [In] object sheet, [In] object success)
        {
            if (!Validate("WorkbookRowsetComplete"))
            {
                Invoker.ReleaseParamsArray(wb, description, sheet, success);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            string newDescription = ToString(description);
            string newSheet = ToString(sheet);
            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newDescription;
            paramsArray[2] = newSheet;
            paramsArray[3] = newSuccess;
            EventBinding.RaiseCustomEvent("WorkbookRowsetComplete", ref paramsArray);
        }
       
        public void AfterCalculate()
        {
            if (!Validate("AfterCalculate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("AfterCalculate", ref paramsArray);
        }
        
        public void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
        {
            if (!Validate("SheetPivotTableAfterValueChange"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, targetRange);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Range newTargetRange = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, targetRange, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newTargetRange;
            EventBinding.RaiseCustomEvent("SheetPivotTableAfterValueChange", ref paramsArray);
        }
        
        public void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetPivotTableBeforeAllocateChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            Int32 newValueChangeStart = ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = ToInt32(valueChangeEnd);
            object[] paramsArray = new object[5];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeAllocateChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        public void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetPivotTableBeforeCommitChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[5];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeCommitChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        public void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd)
        {
            if (!Validate("SheetPivotTableBeforeDiscardChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[4];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeDiscardChanges", ref paramsArray);
        }

        public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowOpen"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowOpen", ref paramsArray);
        }
  
        public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeEdit"))
            {
                Invoker.ReleaseParamsArray(pvw, cancel);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newPvw;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeEdit", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }
        
        public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] object reason, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeClose"))
            {
                Invoker.ReleaseParamsArray(pvw, reason, cancel);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason newReason = (NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason)reason;
            object[] paramsArray = new object[3];
            paramsArray[0] = newPvw;
            paramsArray[1] = newReason;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }
        
        public void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowResize"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowResize", ref paramsArray);
        }
   
        public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowActivate"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowActivate", ref paramsArray);
        }
        
        public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowDeactivate", ref paramsArray);
        }

        public void WorkbookAfterSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object success)
        {
            if (!Validate("WorkbookAfterSave"))
            {
                Invoker.ReleaseParamsArray(wb, success);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newSuccess;
            EventBinding.RaiseCustomEvent("WorkbookAfterSave", ref paramsArray);
        }

        public void WorkbookNewChart([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object ch)
        {
            if (!Validate("WorkbookNewChart"))
            {
                Invoker.ReleaseParamsArray(wb, ch);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Chart newCh = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Chart>(EventClass, ch, NetOffice.ExcelApi.Chart.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newCh;
            EventBinding.RaiseCustomEvent("WorkbookNewChart", ref paramsArray);
        }
     
        public void SheetLensGalleryRenderComplete([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetLensGalleryRenderComplete"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetLensGalleryRenderComplete", ref paramsArray);
        }

        public void SheetTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetTableUpdate"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.TableObject newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.TableObject>(EventClass, target, NetOffice.ExcelApi.TableObject.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetTableUpdate", ref paramsArray);
        }
       
        public void WorkbookModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object changes)
        {
            if (!Validate("WorkbookModelChange"))
            {
                Invoker.ReleaseParamsArray(wb, changes);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
            NetOffice.ExcelApi.ModelChanges newChanges = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ModelChanges>(EventClass, changes, NetOffice.ExcelApi.ModelChanges.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newChanges;
            EventBinding.RaiseCustomEvent("WorkbookModelChange", ref paramsArray);
        }

       
        public void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetBeforeDelete", ref paramsArray);
        }

        #endregion
    }

    #endregion

    #pragma warning restore
}