using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VisioApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[ComImport, Guid("000D0B02-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EWindow
	{
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(701)]
		void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16385)]
		void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4224)]
		void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(702)]
		void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(703)]
		void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(704)]
		void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8193)]
		void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(705)]
		void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(706)]
		void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(707)]
		void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(708)]
		void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(709)]
		void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(710)]
		void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(711)]
		void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(712)]
		void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(713)]
		void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(714)]
		void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EWindow_SinkHelper : SinkHelper, EWindow
	{
		#region Static
		
		public static readonly string Id = "000D0B02-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public EWindow_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region EWindow Members
		
		public void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("SelectionChanged", ref paramsArray);
		}

		public void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeWindowClosed");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("BeforeWindowClosed", ref paramsArray);
		}

		public void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowActivated", ref paramsArray);
		}

		public void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeWindowSelDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("BeforeWindowSelDelete", ref paramsArray);
		}

		public void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeWindowPageTurn");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("BeforeWindowPageTurn", ref paramsArray);
		}

		public void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowTurnedToPage");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowTurnedToPage", ref paramsArray);
		}

		public void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowChanged", ref paramsArray);
		}

		public void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("ViewChanged", ref paramsArray);
		}

		public void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelWindowClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("QueryCancelWindowClose", ref paramsArray);
		}

		public void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowCloseCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowCloseCanceled", ref paramsArray);
		}

		public void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnKeystrokeMessageForAddon");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(mSG);
				return;
			}

			NetOffice.VisioApi.IVMSGWrap newMSG = Factory.CreateObjectFromComProxy(_eventClass, mSG) as NetOffice.VisioApi.IVMSGWrap;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMSG;
			_eventBinding.RaiseCustomEvent("OnKeystrokeMessageForAddon", ref paramsArray);
		}

		public void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			Double newx = Convert.ToDouble(x);
			Double newy = Convert.ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			_eventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			cancelDefault = (bool)paramsArray[4];
		}

		public void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			Double newx = Convert.ToDouble(x);
			Double newy = Convert.ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			_eventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

			cancelDefault = (bool)paramsArray[4];
		}

		public void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			Double newx = Convert.ToDouble(x);
			Double newy = Convert.ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

			cancelDefault = (bool)paramsArray[4];
		}

		public void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			_eventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			cancelDefault = (bool)paramsArray[2];
		}

		public void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii, cancelDefault);
				return;
			}

			Int32 newKeyAscii = Convert.ToInt32(keyAscii);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray.SetValue(cancelDefault, 1);
			_eventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			cancelDefault = (bool)paramsArray[1];
		}

		public void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			_eventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			cancelDefault = (bool)paramsArray[2];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}