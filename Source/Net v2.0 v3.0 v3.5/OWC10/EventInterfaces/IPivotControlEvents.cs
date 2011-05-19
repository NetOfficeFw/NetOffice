using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OWC10Api
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("OWC10", 1)]
	[ComImport, Guid("F5B39A87-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IPivotControlEvents
	{
		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6003)]
		void SelectionChange();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6004)]
		void ViewChange([In] object reason);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6007)]
		void DataChange([In] object reason);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6021)]
		void PivotTableChange([In] object reason);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6043)]
		void BeforeQuery();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6044)]
		void Query();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6029)]
		void OnConnect();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6030)]
		void OnDisconnect();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6034)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6032)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6033)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6035)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6005)]
		void Click();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6006)]
		void DblClick();

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1000)]
		void CommandEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1001)]
		void CommandChecked([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void CommandTipText([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void CommandBeforeExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In] object succeeded);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1009)]
		void KeyDown([In] object keyCode, [In] object shift);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void KeyUp([In] object keyCode, [In] object shift);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1010)]
		void KeyPress([In] object keyAscii);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1006)]
		void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1005)]
		void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1007)]
		void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1011)]
		void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6045)]
		void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object activeObject, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object arrowMode, [In, MarshalAs(UnmanagedType.IDispatch)] object caretPosition, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6046)]
		void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByLibrary("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6049)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class IPivotControlEvents_SinkHelper : SinkHelper, IPivotControlEvents
	{
		#region Static
		
		public static readonly string Id = "F5B39A87-1480-11D3-8549-00C04FAC67D7";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public IPivotControlEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region IPivotControlEvents Members
		
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

		public void ViewChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			NetOffice.OWC10Api.Enums.PivotViewReasonEnum newReason = (NetOffice.OWC10Api.Enums.PivotViewReasonEnum)reason;
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

			NetOffice.OWC10Api.Enums.PivotDataReasonEnum newReason = (NetOffice.OWC10Api.Enums.PivotDataReasonEnum)reason;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
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

			NetOffice.OWC10Api.Enums.PivotTableReasonEnum newReason = (NetOffice.OWC10Api.Enums.PivotTableReasonEnum)reason;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
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

		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
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

		public void DblClick()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DblClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
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
			NetOffice.OWC10Api.ByRef newEnabled = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, enabled) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
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
			NetOffice.OWC10Api.ByRef newChecked = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, _checked) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void CommandTipText([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandTipText");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, caption);
				return;
			}

			object newCommand = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, command) as object;
			NetOffice.OWC10Api.ByRef newCaption = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, caption) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCaption;
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
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void CommandExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In] object succeeded)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, succeeded);
				return;
			}

			object newCommand = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, command) as object;
			bool newSucceeded = (bool)succeeded;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newSucceeded;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void KeyDown([In] object keyCode, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			Int32 newKeyCode = (Int32)keyCode;
			Int32 newShift = (Int32)shift;
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void KeyUp([In] object keyCode, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			Int32 newKeyCode = (Int32)keyCode;
			Int32 newShift = (Int32)shift;
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void KeyPress([In] object keyAscii)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii);
				return;
			}

			Int32 newKeyAscii = (Int32)keyAscii;
			object[] paramsArray = new object[1];
			paramsArray[0] = newKeyAscii;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeKeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift, cancel);
				return;
			}

			Int32 newKeyCode = (Int32)keyCode;
			Int32 newShift = (Int32)shift;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			paramsArray[2] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeKeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift, cancel);
				return;
			}

			Int32 newKeyCode = (Int32)keyCode;
			Int32 newShift = (Int32)shift;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			paramsArray[2] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeKeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii, cancel);
				return;
			}

			Int32 newKeyAscii = (Int32)keyAscii;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray[1] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeContextMenu");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(x, y, menu, cancel);
				return;
			}

			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			NetOffice.OWC10Api.ByRef newMenu = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, menu) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[4];
			paramsArray[0] = newx;
			paramsArray[1] = newy;
			paramsArray[2] = newMenu;
			paramsArray[3] = newCancel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object activeObject, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object arrowMode, [In, MarshalAs(UnmanagedType.IDispatch)] object caretPosition, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StartEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection, activeObject, initialValue, arrowMode, caretPosition, cancel, errorDescription);
				return;
			}

			object newSelection = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, selection) as object;
			object newActiveObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, activeObject) as object;
			NetOffice.OWC10Api.ByRef newInitialValue = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, initialValue) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newArrowMode = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, arrowMode) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCaretPosition = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, caretPosition) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newErrorDescription = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, errorDescription) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[7];
			paramsArray[0] = newSelection;
			paramsArray[1] = newActiveObject;
			paramsArray[2] = newInitialValue;
			paramsArray[3] = newArrowMode;
			paramsArray[4] = newCaretPosition;
			paramsArray[5] = newCancel;
			paramsArray[6] = newErrorDescription;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EndEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(accept, finalValue, cancel, errorDescription);
				return;
			}

			bool newAccept = (bool)accept;
			NetOffice.OWC10Api.ByRef newFinalValue = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, finalValue) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newErrorDescription = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, errorDescription) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[4];
			paramsArray[0] = newAccept;
			paramsArray[1] = newFinalValue;
			paramsArray[2] = newCancel;
			paramsArray[3] = newErrorDescription;
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

			NetOffice.OWC10Api.ByRef newScreenTipText = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, screenTipText) as NetOffice.OWC10Api.ByRef;
			object newSourceObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}