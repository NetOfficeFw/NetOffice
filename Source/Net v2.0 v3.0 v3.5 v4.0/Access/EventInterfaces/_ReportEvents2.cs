using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.AccessApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("Access", 12,14)]
	[ComImport, Guid("D7281A87-4B30-41C5-AB7B-FABF9A35442A"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _ReportEvents2
	{
		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2066)]
		void Open([In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2070)]
		void Close();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2071)]
		void Activate();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2072)]
		void Deactivate();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2083)]
		void Error([In] [Out] ref object dataErr, [In] [Out] ref object response);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2161)]
		void NoData([In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2162)]
		void Page();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2058)]
		void Current();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2067)]
		void Load();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2068)]
		void Resize();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2069)]
		void Unload([In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2084)]
		void Timer();

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2155)]
		void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType);

		[SupportByLibraryAttribute("Access", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2401)]
		void MouseWheel([In] object page, [In] object count);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ReportEvents2_SinkHelper : SinkHelper, _ReportEvents2
	{
		#region Static
		
		public static readonly string Id = "D7281A87-4B30-41C5-AB7B-FABF9A35442A";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _ReportEvents2_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _ReportEvents2 Members
		
		public void Open([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Open");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Close()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Close");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Activate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Activate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Deactivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Deactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Error([In] [Out] ref object dataErr, [In] [Out] ref object response)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Error");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataErr, response);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(dataErr, 0);
			paramsArray.SetValue(response, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			dataErr = (Int16)paramsArray[0];
			response = (Int16)paramsArray[1];
		}

		public void NoData([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NoData");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Page()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Page");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Current()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Current");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Load()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Load");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Resize()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Resize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Unload([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Unload");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			keyCode = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
		}

		public void Timer()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Timer");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Filter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, filterType);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(filterType, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
			filterType = (Int16)paramsArray[1];
		}

		public void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ApplyFilter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, applyType);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(applyType, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
			applyType = (Int16)paramsArray[1];
		}

		public void MouseWheel([In] object page, [In] object count)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseWheel");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page, count);
				return;
			}

			bool newPage = (bool)page;
			Int32 newCount = (Int32)count;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}