using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.EventContracts
{	
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00024412-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface WorkbookEvents
	{
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(682)]
		void Open();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1546)]
		void BeforeClose([In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1547)]
		void BeforeSave([In] object saveAsUI, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1549)]
		void BeforePrint([In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1550)]
		void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1552)]
		void AddinInstall();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1553)]
		void AddinUninstall();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(ExcelApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
		void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
		void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
		void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.Hyperlink))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
		void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2158)]
		void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2159)]
		void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 11,12,14,15,16)]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2266)]
		void Sync([In] object syncEventType);
        
        [SupportByVersion("Excel", 11,12,14,15,16)]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("isRefresh", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2283)]
		void BeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 11,12,14,15,16)]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("isRefresh", SinkArgumentType.Bool)]
        [SinkArgument("result", SinkArgumentType.Enum, typeof(ExcelApi.Enums.XlXmlImportResult))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2285)]
		void AfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result);

		[SupportByVersion("Excel", 11,12,14,15,16)]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2287)]
		void BeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 11,12,14,15,16)]
        [SinkArgument("map", typeof(ExcelApi.XmlMap))]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2288)]
		void AfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result);

		[SupportByVersion("Excel", 12,14,15,16)]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("sheet", SinkArgumentType.String)]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2610)]
		void RowsetComplete([In] object description, [In] object sheet, [In] object success);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("targetRange", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
		void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.Range))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
		void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.Range))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
		void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.Range))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
		void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2899)]
		void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2900)]
		void AfterSave([In] object success);

		[SupportByVersion("Excel", 14,15,16)]
        [SinkArgument("ch", typeof(ExcelApi.Chart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2901)]
		void NewChart([In, MarshalAs(UnmanagedType.IDispatch)] object ch);

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
        [SinkArgument("changes", typeof(ExcelApi.ModelChanges))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3075)]
		void ModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object changes);

		[SupportByVersion("Excel", 15, 16)]
        [SinkArgument("sh", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3077)]
		void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
	}
}