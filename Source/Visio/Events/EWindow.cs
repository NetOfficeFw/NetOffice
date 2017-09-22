using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B02-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EWindow
	{
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(701)]
		void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16385)]
		void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4224)]
		void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(702)]
		void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(703)]
		void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(704)]
		void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8193)]
		void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(705)]
		void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(706)]
		void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(NetOffice.VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(707)]
		void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("mSG", typeof(NetOffice.VisioApi.IVMSGWrap))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(708)]
		void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(709)]
		void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(710)]
		void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(711)]
		void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(712)]
		void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(713)]
		void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(714)]
		void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EWindow_SinkHelper : SinkHelper, EWindow
	{
		#region Static
		
		public static readonly string Id = "000D0B02-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public EWindow_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

        #endregion

        #region EWindow

        public void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("SelectionChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("SelectionChanged", ref paramsArray);
        }

        public void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowClosed"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowClosed", ref paramsArray);
        }

        public void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowActivated"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowActivated", ref paramsArray);
        }

        public void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowSelDelete"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowSelDelete", ref paramsArray);
        }

        public void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowPageTurn"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowPageTurn", ref paramsArray);
        }

        public void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowTurnedToPage"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowTurnedToPage", ref paramsArray);
        }

        public void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
            if (!Validate("WindowChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			EventBinding.RaiseCustomEvent("WindowChanged", ref paramsArray);
		}

        public void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
            if (!Validate("ViewChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			EventBinding.RaiseCustomEvent("ViewChanged", ref paramsArray);
		}

        public void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
            if (!Validate("QueryCancelWindowClose"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			EventBinding.RaiseCustomEvent("QueryCancelWindowClose", ref paramsArray);
		}

        public void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
            if (!Validate("WindowCloseCanceled"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			EventBinding.RaiseCustomEvent("WindowCloseCanceled", ref paramsArray);
		}

        public void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG)
        {
            if (!Validate("OnKeystrokeMessageForAddon"))
            {
                Invoker.ReleaseParamsArray(mSG);
                return;
            }

            NetOffice.VisioApi.IVMSGWrap newMSG = Factory.CreateEventArgumentObjectFromComProxy(EventClass, mSG) as NetOffice.VisioApi.IVMSGWrap;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMSG;
            EventBinding.RaiseCustomEvent("OnKeystrokeMessageForAddon", ref paramsArray);
        }

        public void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
            Int32 newKeyButtonState = ToInt32(keyButtonState);
            Double newx = ToDouble(x);
            Double newy = ToDouble(y);
            object[] paramsArray = new object[5];
            paramsArray[0] = newButton;
            paramsArray[1] = newKeyButtonState;
            paramsArray[2] = newx;
            paramsArray[3] = newy;
            paramsArray.SetValue(cancelDefault, 4);
            EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[4]);
        }

        public void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
            Int32 newKeyButtonState = ToInt32(keyButtonState);
            Double newx = ToDouble(x);
            Double newy = ToDouble(y);
            object[] paramsArray = new object[5];
            paramsArray[0] = newButton;
            paramsArray[1] = newKeyButtonState;
            paramsArray[2] = newx;
            paramsArray[3] = newy;
            paramsArray.SetValue(cancelDefault, 4);
            EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[4]);
        }

        public void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
            Int32 newKeyButtonState = ToInt32(keyButtonState);
            Double newx = ToDouble(x);
            Double newy = ToDouble(y);
            object[] paramsArray = new object[5];
            paramsArray[0] = newButton;
            paramsArray[1] = newKeyButtonState;
            paramsArray[2] = newx;
            paramsArray[3] = newy;
            paramsArray.SetValue(cancelDefault, 4);
            EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[4]);
        }

        public void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
                return;
            }

            Int32 newKeyCode = ToInt32(keyCode);
            Int32 newKeyButtonState = ToInt32(keyButtonState);
            object[] paramsArray = new object[3];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newKeyButtonState;
            paramsArray.SetValue(cancelDefault, 2);
            EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[2]);
        }

        public void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii, cancelDefault);
                return;
            }

            Int32 newKeyAscii = ToInt32(keyAscii);
            object[] paramsArray = new object[2];
            paramsArray[0] = newKeyAscii;
            paramsArray.SetValue(cancelDefault, 1);
            EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[1]);
        }

        public void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
                return;
            }

            Int32 newKeyCode = ToInt32(keyCode);
            Int32 newKeyButtonState = ToInt32(keyButtonState);
            object[] paramsArray = new object[3];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newKeyButtonState;
            paramsArray.SetValue(cancelDefault, 2);
            EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[2]);
        }

        #endregion
    }

    #endregion

    #pragma warning restore
}