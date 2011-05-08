using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.AccessApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("AC12","AC14")]
	[ComImport, Guid("0EA530DD-5B30-4278-BD28-47C4D11619BD"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _FormEvents2
	{
		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2067)]
		void Load();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2058)]
		void Current();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2059)]
		void BeforeInsert([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2060)]
		void AfterInsert();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2063)]
		void Delete([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2064)]
		void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2065)]
		void AfterDelConfirm([In] [Out] ref object status);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2066)]
		void Open([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2068)]
		void Resize();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2069)]
		void Unload([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2070)]
		void Close();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2071)]
		void Activate();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2072)]
		void Deactivate();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2083)]
		void Error([In] [Out] ref object dataErr, [In] [Out] ref object response);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2084)]
		void Timer();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2155)]
		void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2145)]
		void Undo([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2334)]
		void RecordExit([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2369)]
		void BeginBatchEdit([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2370)]
		void UndoBatchEdit([In] [Out] ref object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2383)]
		void OnConnect();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2384)]
		void OnDisconnect();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2385)]
		void PivotTableChange([In] object reason);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2386)]
		void Query();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2387)]
		void BeforeQuery();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2388)]
		void SelectionChange();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2389)]
		void CommandBeforeExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2390)]
		void CommandChecked([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2391)]
		void CommandEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2392)]
		void CommandExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2394)]
		void DataSetChange();

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2395)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2399)]
		void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2397)]
		void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2396)]
		void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2398)]
		void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2401)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2402)]
		void ViewChange([In] object reason);

		[SupportByLibrary("AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2403)]
		void DataChange([In] object reason);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _FormEvents2_SinkHelper : SinkHelper, _FormEvents2
	{
		#region Static
		
		public static readonly string Id = "0EA530DD-5B30-4278-BD28-47C4D11619BD";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _FormEvents2_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _FormEvents2 Members
		
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

		public void BeforeInsert([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeInsert");
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

		public void AfterInsert()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Delete([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Delete");
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

		public void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDelConfirm");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, response);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(response, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
			response = (Int16)paramsArray[1];
		}

		public void AfterDelConfirm([In] [Out] ref object status)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterDelConfirm");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(status);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(status, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			status = (Int16)paramsArray[0];
		}

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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Undo([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Undo");
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

		public void RecordExit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RecordExit");
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

		public void BeginBatchEdit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeginBatchEdit");
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

		public void UndoBatchEdit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("UndoBatchEdit");
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

		public void OnConnect()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnConnect");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void OnDisconnect()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnDisconnect");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			Int32 newReason = (Int32)reason;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Query()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Query");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeQuery()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeQuery");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SelectionChange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void CommandBeforeExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandBeforeExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, cancel);
				return;
			}

			object newCommand = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, command) as object;
			object newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void CommandChecked([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandChecked");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, _checked);
				return;
			}

			object newCommand = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, command) as object;
			object newChecked = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, _checked) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void CommandEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandEnabled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, enabled);
				return;
			}

			object newCommand = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, command) as object;
			object newEnabled = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, enabled) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void CommandExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command);
				return;
			}

			object newCommand = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, command) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCommand;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DataSetChange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataSetChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeScreenTip");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(screenTipText, sourceObject);
				return;
			}

			object newScreenTipText = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, screenTipText) as object;
			object newSourceObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeRender");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject, chartObject, cancel);
				return;
			}

			object newdrawObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object newchartObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, chartObject) as object;
			object newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			paramsArray[2] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterRender");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject, chartObject);
				return;
			}

			object newdrawObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object newchartObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, chartObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterFinalRender");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject);
				return;
			}

			object newdrawObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterLayout");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject);
				return;
			}

			object newdrawObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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

		public void ViewChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			Int32 newReason = (Int32)reason;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DataChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			Int32 newReason = (Int32)reason;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}