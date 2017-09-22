using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F5B39A7A-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IChartEvents
	{
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5101)]
		void DataSetChange();

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5102)]
		void DblClick();

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5103)]
		void Click();

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1009)]
		void KeyDown([In] object keyCode, [In] object shift);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void KeyUp([In] object keyCode, [In] object shift);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1010)]
		void KeyPress([In] object keyAscii);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1006)]
		void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1005)]
		void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1007)]
		void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5107)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5108)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5109)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("page", SinkArgumentType.Bool)]
        [SinkArgument("count", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5118)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5110)]
		void SelectionChange();

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("tipText", typeof(OWC10Api.ByRef))]
        [SinkArgument("newContextObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5120)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object tipText, [In, MarshalAs(UnmanagedType.IDispatch)] object contextObject);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("enabled", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1000)]
		void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("_checked", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1001)]
		void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("caption", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void CommandTipText([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("succeeded", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In] object command, [In] object succeeded);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [SinkArgument("menu", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1011)]
		void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5111)]
		void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [SinkArgument("chartObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5112)]
		void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5113)]
		void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("drawObject", typeof(OWC10Api.ChChartDraw))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5114)]
		void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5119)]
		void ViewChange();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class IChartEvents_SinkHelper : SinkHelper, IChartEvents
	{
		#region Static
		
		public static readonly string Id = "F5B39A7A-1480-11D3-8549-00C04FAC67D7";
		
		#endregion

		#region Ctor

		public IChartEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region IChartEvents
		
		public void DataSetChange()
		{
            if (!Validate("DataSetChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DataSetChange", ref paramsArray);
		}

		public void DblClick()
		{
            if (!Validate("DblClick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
		}

		public void Click()
		{
            if (!Validate("Click"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

        public void KeyDown([In] object keyCode, [In] object shift)
		{
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			Int32 newKeyCode = ToInt32(keyCode);
			Int32 newShift = ToInt32(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);
		}

        public void KeyUp([In] object keyCode, [In] object shift)
		{
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

            Int32 newKeyCode = ToInt32(keyCode);
			Int32 newShift = ToInt32(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
		}

        public void KeyPress([In] object keyAscii)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

			Int32 newKeyAscii = ToInt32(keyAscii);
			object[] paramsArray = new object[1];
			paramsArray[0] = newKeyAscii;
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);
		}

        public void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
            if (!Validate("BeforeKeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift, cancel);
                return;
            }

			Int32 newKeyCode = ToInt32(keyCode);
			Int32 newShift = ToInt32(shift);
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			paramsArray[2] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeKeyDown", ref paramsArray);
		}

        public void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeKeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift, cancel);
                return;
            }

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newShift = Convert.ToInt32(shift);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			paramsArray[2] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeKeyUp", ref paramsArray);
		}

        public void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
            if (!Validate("BeforeKeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii, cancel);
                return;
            }

			Int32 newKeyAscii = Convert.ToInt32(keyAscii);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray[1] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeKeyPress", ref paramsArray);
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

        public void MouseWheel([In] object page, [In] object count)
		{
            if (!Validate("MouseWheel"))
            {
                Invoker.ReleaseParamsArray(page, count);
                return;
            }

			bool newPage = ToBoolean(page);
			Int32 newCount = ToInt32(count);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			EventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
		}

		public void SelectionChange()
		{
            if (!Validate("SelectionChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
		}

        public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object tipText, [In, MarshalAs(UnmanagedType.IDispatch)] object contextObject)
        {
            if (!Validate("BeforeScreenTip"))
            {
                Invoker.ReleaseParamsArray(tipText, contextObject);
                return;
            }

			NetOffice.OWC10Api.ByRef newTipText = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, tipText, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			object newContextObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, contextObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTipText;
			paramsArray[1] = newContextObject;
			EventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
		}

        public void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
        {
            if (!Validate("CommandEnabled"))
            {
                Invoker.ReleaseParamsArray(command, enabled);
                return;
            }

			object newCommand = (object)command;
            NetOffice.OWC10Api.ByRef newEnabled = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, enabled, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
			EventBinding.RaiseCustomEvent("CommandEnabled", ref paramsArray);
        }

        public void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked)
        {
            if (!Validate("CommandChecked"))
            {
                Invoker.ReleaseParamsArray(command, _checked);
                return;
            }

			object newCommand = (object)command;
            NetOffice.OWC10Api.ByRef newChecked = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, _checked, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			EventBinding.RaiseCustomEvent("CommandChecked", ref paramsArray);
		}

        public void CommandTipText([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption)
        {
            if (!Validate("CommandTipText"))
            {
                Invoker.ReleaseParamsArray(command, caption);
                return;
            }
            
			object newCommand = (object)command;
            NetOffice.OWC10Api.ByRef newCaption = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, command, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCaption;
			EventBinding.RaiseCustomEvent("CommandTipText", ref paramsArray);
		}

        public void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("CommandBeforeExecute"))
            {
                Invoker.ReleaseParamsArray(command, cancel);
                return;
            }

			object newCommand = (object)command;
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			EventBinding.RaiseCustomEvent("CommandBeforeExecute", ref paramsArray);
		}

        public void CommandExecute([In] object command, [In] object succeeded)
        {
            if (!Validate("CommandExecute"))
            {
                Invoker.ReleaseParamsArray(command, succeeded);
                return;
            }

			object newCommand = (object)command;
			bool newSucceeded = ToBoolean(succeeded);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newSucceeded;
			EventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
		}

        public void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeContextMenu"))
            {
                Invoker.ReleaseParamsArray(x, y, menu, cancel);
                return;
            }

			Int32 newx = ToInt32(x);
			Int32 newy = ToInt32(y);
            NetOffice.OWC10Api.ByRef newMenu = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, menu, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[4];
			paramsArray[0] = newx;
			paramsArray[1] = newy;
			paramsArray[2] = newMenu;
			paramsArray[3] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeContextMenu", ref paramsArray);
		}

        public void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
            if (!Validate("BeforeRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject, cancel);
                return;
            }

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, NetOffice.OWC10Api.ChChartDraw.LateBindingApiWrapperType);
			object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			paramsArray[2] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeRender", ref paramsArray);
		}

        public void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject)
        {
            if (!Validate("AfterRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject);
                return;
            }

			NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, NetOffice.OWC10Api.ChChartDraw.LateBindingApiWrapperType);
			object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			EventBinding.RaiseCustomEvent("AfterRender", ref paramsArray);
		}

        public void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
		{
            if (!Validate("AfterFinalRender"))
            {
                Invoker.ReleaseParamsArray(drawObject);
                return;
            }

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, NetOffice.OWC10Api.ChChartDraw.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			EventBinding.RaiseCustomEvent("AfterFinalRender", ref paramsArray);
		}

        public void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
        {
            if (!Validate("AfterLayout"))
            {
                Invoker.ReleaseParamsArray(drawObject);
                return;
            }

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, NetOffice.OWC10Api.ChChartDraw.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			EventBinding.RaiseCustomEvent("AfterLayout", ref paramsArray);
		}

		public void ViewChange()
        {
            if (!Validate("ViewChange"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}