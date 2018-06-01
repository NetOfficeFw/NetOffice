using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.EventContracts
{
    /// <summary>
    /// ChartEvents
    /// </summary>
    [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0002440F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ChartEvents
	{
        /// <summary>
        /// Activate
        /// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

        /// <summary>
        /// Deactivate
        /// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

        /// <summary>
        /// Resize
        /// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(256)]
		void Resize();

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1531)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1532)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1533)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// BeforeRightClick
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In] [Out] ref object cancel);

        /// <summary>
        /// DragPlot
        /// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1535)]
		void DragPlot();

        /// <summary>
        /// DragOver
        /// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1536)]
		void DragOver();

        /// <summary>
        /// BeforeDoubleClick
        /// </summary>
        /// <param name="elementID"></param>
        /// <param name="arg1"></param>
        /// <param name="arg2"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("elementID", SinkArgumentType.Int32)]
        [SinkArgument("arg1", SinkArgumentType.Int32)]
        [SinkArgument("arg2", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In] object elementID, [In] object arg1, [In] object arg2, [In] [Out] ref object cancel);

        /// <summary>
        /// Select
        /// </summary>
        /// <param name="elementID"></param>
        /// <param name="arg1"></param>
        /// <param name="arg2"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("elementID", SinkArgumentType.Int32)]
        [SinkArgument("arg1", SinkArgumentType.Int32)]
        [SinkArgument("arg2", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(235)]
		void Select([In] object elementID, [In] object arg1, [In] object arg2);

        /// <summary>
        /// SeriesChange
        /// </summary>
        /// <param name="seriesIndex"></param>
        /// <param name="pointIndex"></param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("seriesIndex", SinkArgumentType.Int32)]
        [SinkArgument("pointIndex", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1538)]
		void SeriesChange([In] object seriesIndex, [In] object pointIndex);

        /// <summary>
        /// Calculate
        /// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();
	}
}
