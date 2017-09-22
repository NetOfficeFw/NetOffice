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
    [ComImport, Guid("F5B39A87-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IPivotControlEvents
	{
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6003)]
		void SelectionChange();

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("reason", SinkArgumentType.Enum, typeof(OWC10Api.Enums.PivotViewReasonEnum))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6004)]
		void ViewChange([In] object reason);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("reason", SinkArgumentType.Enum, typeof(OWC10Api.Enums.PivotViewReasonEnum))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6007)]
		void DataChange([In] object reason);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("reason", SinkArgumentType.Enum, typeof(OWC10Api.Enums.PivotViewReasonEnum))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6021)]
		void PivotTableChange([In] object reason);

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6043)]
		void BeforeQuery();

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6044)]
		void Query();

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6029)]
		void OnConnect();

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6030)]
		void OnDisconnect();

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6034)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6032)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6033)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("page", SinkArgumentType.Bool)]
        [SinkArgument("count", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6035)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6005)]
		void Click();

		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6006)]
		void DblClick();

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
        [SinkArgument("succeeded", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In] object command, [In] object succeeded);

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
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [SinkArgument("menu", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1011)]
		void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("selection", SinkArgumentType.UnknownProxy)]
        [SinkArgument("activeObject", SinkArgumentType.UnknownProxy)]
        [SinkArgument("initialValue", typeof(OWC10Api.ByRef))]
        [SinkArgument("arrowMode", typeof(OWC10Api.ByRef))]
        [SinkArgument("caretPosition", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [SinkArgument("errorDescription", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6045)]
		void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object activeObject, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object arrowMode, [In, MarshalAs(UnmanagedType.IDispatch)] object caretPosition, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("accept", SinkArgumentType.Bool)]
        [SinkArgument("finalValue", typeof(OWC10Api.ByRef))]
        [SinkArgument("cancel", typeof(OWC10Api.ByRef))]
        [SinkArgument("errorDescription", typeof(OWC10Api.ByRef))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6046)]
		void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("screenTipText", typeof(OWC10Api.ByRef))]
        [SinkArgument("sourceObject", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6049)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class IPivotControlEvents_SinkHelper : SinkHelper, IPivotControlEvents
    {
        #region Static

        public static readonly string Id = "F5B39A87-1480-11D3-8549-00C04FAC67D7";

        #endregion

        #region Ctor

        public IPivotControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region IPivotControlEvents

        public void SelectionChange()
        {
            if (!Validate("SelectionChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
        }

        public void ViewChange([In] object reason)
        {
            if (!Validate("ViewChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

            NetOffice.OWC10Api.Enums.PivotViewReasonEnum newReason = (NetOffice.OWC10Api.Enums.PivotViewReasonEnum)reason;
            object[] paramsArray = new object[1];
            paramsArray[0] = newReason;
            EventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
        }

        public void DataChange([In] object reason)
        {
            if (!Validate("DataChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

            NetOffice.OWC10Api.Enums.PivotDataReasonEnum newReason = (NetOffice.OWC10Api.Enums.PivotDataReasonEnum)reason;
            object[] paramsArray = new object[1];
            paramsArray[0] = newReason;
            EventBinding.RaiseCustomEvent("DataChange", ref paramsArray);
        }

        public void PivotTableChange([In] object reason)
        {
            if (!Validate("PivotTableChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

            NetOffice.OWC10Api.Enums.PivotTableReasonEnum newReason = (NetOffice.OWC10Api.Enums.PivotTableReasonEnum)reason;
            object[] paramsArray = new object[1];
            paramsArray[0] = newReason;
            EventBinding.RaiseCustomEvent("PivotTableChange", ref paramsArray);
        }

        public void BeforeQuery()
        {
            if (!Validate("BeforeQuery"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("BeforeQuery", ref paramsArray);
        }

        public void Query()
        {
            if (!Validate("Query"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Query", ref paramsArray);
        }

        public void OnConnect()
        {
            if (!Validate("OnConnect"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("OnConnect", ref paramsArray);
        }

        public void OnDisconnect()
        {
            if (!Validate("OnDisconnect"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("OnDisconnect", ref paramsArray);
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

            bool newPage = Convert.ToBoolean(page);
            Int32 newCount = Convert.ToInt32(count);
            object[] paramsArray = new object[2];
            paramsArray[0] = newPage;
            paramsArray[1] = newCount;
            EventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
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

        public void DblClick()
        {
            if (!Validate("DblClick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
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

        public void CommandExecute([In] object command, [In] object succeeded)
        {
            if (!Validate("CommandExecute"))
            {
                Invoker.ReleaseParamsArray(command, succeeded);
                return;
            }

            object newCommand = (object)command;
            bool newSucceeded = Convert.ToBoolean(succeeded);
            object[] paramsArray = new object[2];
            paramsArray[0] = newCommand;
            paramsArray[1] = newSucceeded;
            EventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
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

            Int32 newKeyAscii = Convert.ToInt32(keyAscii);
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

            Int32 newKeyCode = ToInt32(keyCode);
            Int32 newShift = ToInt32(shift);
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

            Int32 newKeyAscii = ToInt32(keyAscii);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newKeyAscii;
            paramsArray[1] = newCancel;
            EventBinding.RaiseCustomEvent("BeforeKeyPress", ref paramsArray);
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

        public void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object activeObject, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object arrowMode, [In, MarshalAs(UnmanagedType.IDispatch)] object caretPosition, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
        {
            if (!Validate("StartEdit"))
            {
                Invoker.ReleaseParamsArray(selection, activeObject, initialValue, arrowMode, caretPosition, cancel, errorDescription);
                return;
            }

			object newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as object;
			object newActiveObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, activeObject) as object;
			NetOffice.OWC10Api.ByRef newInitialValue = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, initialValue, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			NetOffice.OWC10Api.ByRef newArrowMode = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, arrowMode, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			NetOffice.OWC10Api.ByRef newCaretPosition = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, caretPosition, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			NetOffice.OWC10Api.ByRef newErrorDescription = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, errorDescription, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			object[] paramsArray = new object[7];
			paramsArray[0] = newSelection;
			paramsArray[1] = newActiveObject;
			paramsArray[2] = newInitialValue;
			paramsArray[3] = newArrowMode;
			paramsArray[4] = newCaretPosition;
			paramsArray[5] = newCancel;
			paramsArray[6] = newErrorDescription;
			EventBinding.RaiseCustomEvent("StartEdit", ref paramsArray);
		}

        public void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
		{
            if (!Validate("EndEdit"))
            {
                Invoker.ReleaseParamsArray(accept, finalValue, cancel, errorDescription);
                return;
            }

			bool newAccept = ToBoolean(accept);
			NetOffice.OWC10Api.ByRef newFinalValue = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, finalValue, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			NetOffice.OWC10Api.ByRef newErrorDescription = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, errorDescription, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			object[] paramsArray = new object[4];
			paramsArray[0] = newAccept;
			paramsArray[1] = newFinalValue;
			paramsArray[2] = newCancel;
			paramsArray[3] = newErrorDescription;
			EventBinding.RaiseCustomEvent("EndEdit", ref paramsArray);
		}

        public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject)
		{
            if (!Validate("BeforeScreenTip"))
            {
                Invoker.ReleaseParamsArray(screenTipText, sourceObject);
                return;
            }

			NetOffice.OWC10Api.ByRef newScreenTipText = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, screenTipText, NetOffice.OWC10Api.ByRef.LateBindingApiWrapperType);
			object newSourceObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			EventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}