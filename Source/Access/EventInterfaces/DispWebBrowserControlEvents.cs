using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.AccessApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Access", 14,15,16)]
	[ComImport, Guid("EACB9075-68F8-4E3B-B865-E1CE6BE0447C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DispWebBrowserControlEvents
	{
		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2076)]
		void Updated([In] [Out] ref object code);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void Enter();

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2075)]
		void Exit([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2077)]
		void Change();

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2524)]
		void BeforeNavigate2([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object flags, [In] [Out] ref object targetFrameName, [In] [Out] ref object postData, [In] [Out] ref object headers, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2528)]
		void DocumentComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2515)]
		void ProgressChange([In] object progress, [In] object progressMax);

		[SupportByVersionAttribute("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2510)]
		void NavigateError([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object targetFrameName, [In] [Out] ref object statusCode, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispWebBrowserControlEvents_SinkHelper : SinkHelper, DispWebBrowserControlEvents
	{
		#region Static
		
		public static readonly string Id = "EACB9075-68F8-4E3B-B865-E1CE6BE0447C";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public DispWebBrowserControlEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region DispWebBrowserControlEvents Members
		
		public void Updated([In] [Out] ref object code)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Updated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(code);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(code, 0);
			_eventBinding.RaiseCustomEvent("Updated", ref paramsArray);

			code = (Int16)paramsArray[0];
		}

		public void BeforeUpdate([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void AfterUpdate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		public void Enter()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Enter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Enter", ref paramsArray);
		}

		public void Exit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Exit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Exit", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Dirty([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Dirty");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Dirty", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Change()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Change");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		public void GotFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GotFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
		}

		public void LostFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LostFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
		}

		public void Click()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Click");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		public void DblClick([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DblClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("DblClick", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			_eventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			button = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
			x = (Single)paramsArray[2];
			y = (Single)paramsArray[3];
		}

		public void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			_eventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

			button = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
			x = (Single)paramsArray[2];
			y = (Single)paramsArray[3];
		}

		public void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

			button = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
			x = (Single)paramsArray[2];
			y = (Single)paramsArray[3];
		}

		public void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			_eventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			keyCode = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
		}

		public void KeyPress([In] [Out] ref object keyAscii)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(keyAscii, 0);
			_eventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			keyAscii = (Int16)paramsArray[0];
		}

		public void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			_eventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			keyCode = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
		}

		public void BeforeNavigate2([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object flags, [In] [Out] ref object targetFrameName, [In] [Out] ref object postData, [In] [Out] ref object headers, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeNavigate2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pDisp, uRL, flags, targetFrameName, postData, headers, cancel);
				return;
			}

			object newpDisp = Factory.CreateObjectFromComProxy(_eventClass, pDisp) as object;
			object[] paramsArray = new object[7];
			paramsArray[0] = newpDisp;
			paramsArray.SetValue(uRL, 1);
			paramsArray.SetValue(flags, 2);
			paramsArray.SetValue(targetFrameName, 3);
			paramsArray.SetValue(postData, 4);
			paramsArray.SetValue(headers, 5);
			paramsArray.SetValue(cancel, 6);
			_eventBinding.RaiseCustomEvent("BeforeNavigate2", ref paramsArray);

			uRL = (object)paramsArray[1];
			flags = (object)paramsArray[2];
			targetFrameName = (object)paramsArray[3];
			postData = (object)paramsArray[4];
			headers = (object)paramsArray[5];
			cancel = (bool)paramsArray[6];
		}

		public void DocumentComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pDisp, uRL);
				return;
			}

			object newpDisp = Factory.CreateObjectFromComProxy(_eventClass, pDisp) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpDisp;
			paramsArray.SetValue(uRL, 1);
			_eventBinding.RaiseCustomEvent("DocumentComplete", ref paramsArray);

			uRL = (object)paramsArray[1];
		}

		public void ProgressChange([In] object progress, [In] object progressMax)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProgressChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(progress, progressMax);
				return;
			}

			Int32 newProgress = Convert.ToInt32(progress);
			Int32 newProgressMax = Convert.ToInt32(progressMax);
			object[] paramsArray = new object[2];
			paramsArray[0] = newProgress;
			paramsArray[1] = newProgressMax;
			_eventBinding.RaiseCustomEvent("ProgressChange", ref paramsArray);
		}

		public void NavigateError([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object targetFrameName, [In] [Out] ref object statusCode, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NavigateError");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pDisp, uRL, targetFrameName, statusCode, cancel);
				return;
			}

			object newpDisp = Factory.CreateObjectFromComProxy(_eventClass, pDisp) as object;
			object[] paramsArray = new object[5];
			paramsArray[0] = newpDisp;
			paramsArray.SetValue(uRL, 1);
			paramsArray.SetValue(targetFrameName, 2);
			paramsArray.SetValue(statusCode, 3);
			paramsArray.SetValue(cancel, 4);
			_eventBinding.RaiseCustomEvent("NavigateError", ref paramsArray);

			uRL = (object)paramsArray[1];
			targetFrameName = (object)paramsArray[2];
			statusCode = (object)paramsArray[3];
			cancel = (bool)paramsArray[4];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}