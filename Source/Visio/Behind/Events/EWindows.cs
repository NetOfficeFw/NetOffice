using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.VisioApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EWindows"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EWindows_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EWindows
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EWindows
		/// </summary>
		public static readonly string Id = "000D0B01-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EWindows_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EWindows

        /// <summary>
        /// 
        /// </summary>
        /// <param name="window"></param>
        public void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
            if (!Validate("WindowOpened"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			EventBinding.RaiseCustomEvent("WindowOpened", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="window"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="mSG"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="cancelDefault"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="cancelDefault"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="cancelDefault"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="cancelDefault"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyAscii"></param>
		/// <param name="cancelDefault"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="keyButtonState"></param>
		/// <param name="cancelDefault"></param>
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
	
}
