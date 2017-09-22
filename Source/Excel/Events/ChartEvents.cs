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

	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0002440F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ChartEvents
	{
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(256)]
		void Resize();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1531)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1532)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1533)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1535)]
		void DragPlot();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1536)]
		void DragOver();

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("elementID", SinkArgumentType.Int32)]
        [SinkArgument("arg1", SinkArgumentType.Int32)]
        [SinkArgument("arg2", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In] object elementID, [In] object arg1, [In] object arg2, [In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("elementID", SinkArgumentType.Int32)]
        [SinkArgument("arg1", SinkArgumentType.Int32)]
        [SinkArgument("arg2", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(235)]
		void Select([In] object elementID, [In] object arg1, [In] object arg2);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("seriesIndex", SinkArgumentType.Int32)]
        [SinkArgument("pointIndex", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1538)]
		void SeriesChange([In] object seriesIndex, [In] object pointIndex);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ChartEvents_SinkHelper : SinkHelper, ChartEvents
	{
		#region Static
		
		public static readonly string Id = "0002440F-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public ChartEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ChartEvents
		
		public void Activate()
		{
            if (!Validate("Activate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		public void Deactivate()
		{
            if (!Validate("Deactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		public void Resize()
		{
            if (!Validate("Resize"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Resize", ref paramsArray);
		}

		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }
            
			Int32 newButton = ToInt32(button);
			Int32 newShift = ToInt32(shift);
			Int32 newx = ToInt32(x);
			Int32 newy = ToInt32(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

		public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			Int32 newButton = ToInt32(button);
			Int32 newShift = ToInt32(shift);
			Int32 newx = ToInt32(x);
			Int32 newy = ToInt32(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			Int32 newButton = ToInt32(button);
			Int32 newShift = ToInt32(shift);
			Int32 newx = ToInt32(x);
			Int32 newy = ToInt32(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

		public void BeforeRightClick([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
		}

		public void DragPlot()
        {
            if (!Validate("DragPlot"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DragPlot", ref paramsArray);
		}

		public void DragOver()
		{
            if (!Validate("DragOver"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DragOver", ref paramsArray);
		}

		public void BeforeDoubleClick([In] object elementID, [In] object arg1, [In] object arg2, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(elementID, arg1, arg2, cancel);
                return;
            }

			Int32 newElementID = ToInt32(elementID);
			Int32 newArg1 = ToInt32(arg1);
			Int32 newArg2 = ToInt32(arg2);
			object[] paramsArray = new object[4];
			paramsArray[0] = newElementID;
			paramsArray[1] = newArg1;
			paramsArray[2] = newArg2;
			paramsArray.SetValue(cancel, 3);
			EventBinding.RaiseCustomEvent("BeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
		}

		public void Select([In] object elementID, [In] object arg1, [In] object arg2)
		{
            if (!Validate("Select"))
            {
                Invoker.ReleaseParamsArray(elementID, arg1, arg2);
                return;
            }

			Int32 newElementID = ToInt32(elementID);
			Int32 newArg1 = ToInt32(arg1);
			Int32 newArg2 = ToInt32(arg2);
			object[] paramsArray = new object[3];
			paramsArray[0] = newElementID;
			paramsArray[1] = newArg1;
			paramsArray[2] = newArg2;
			EventBinding.RaiseCustomEvent("Select", ref paramsArray);
		}

		public void SeriesChange([In] object seriesIndex, [In] object pointIndex)
		{
            if (!Validate("SeriesChange"))
            {
                Invoker.ReleaseParamsArray(seriesIndex, pointIndex);
                return;
            }

			Int32 newSeriesIndex = ToInt32(seriesIndex);
			Int32 newPointIndex = ToInt32(pointIndex);
			object[] paramsArray = new object[2];
			paramsArray[0] = newSeriesIndex;
			paramsArray[1] = newPointIndex;
			EventBinding.RaiseCustomEvent("SeriesChange", ref paramsArray);
		}

		public void Calculate()
		{
            if (!Validate("Calculate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Calculate", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}