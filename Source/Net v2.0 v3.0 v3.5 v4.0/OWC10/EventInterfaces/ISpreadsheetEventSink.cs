using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OWC10Api
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("OWC10", 1)]
	[ComImport, Guid("F5B39A75-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ISpreadsheetEventSink
	{
		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1011)]
		void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1006)]
		void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1007)]
		void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1005)]
		void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1502)]
		void Click();

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1000)]
		void CommandEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1001)]
		void CommandChecked([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void CommandTipText([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void CommandBeforeExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In, MarshalAs(UnmanagedType.IDispatch)] object command, [In] object succeeded);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1503)]
		void DblClick();

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1504)]
		void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1523)]
		void Initialize();

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1009)]
		void KeyDown([In] object keyCode, [In] object shift);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1010)]
		void KeyPress([In] object keyAscii);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void KeyUp([In] object keyCode, [In] object shift);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1522)]
		void LoadCompleted();

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1505)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1506)]
		void MouseOut([In] object button, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1507)]
		void MouseOver([In] object button, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1508)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1509)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1511)]
		void SelectionChange();

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1512)]
		void SelectionChanging([In, MarshalAs(UnmanagedType.IDispatch)] object range);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1513)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1516)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1517)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1518)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1519)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1520)]
		void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByLibraryAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1521)]
		void ViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object target);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ISpreadsheetEventSink_SinkHelper : SinkHelper, ISpreadsheetEventSink
	{
		#region Static
		
		public static readonly string Id = "F5B39A75-1480-11D3-8549-00C04FAC67D7";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ISpreadsheetEventSink_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ISpreadsheetEventSink Members
		
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

		public void Initialize()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Initialize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
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

		public void LoadCompleted()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LoadCompleted");
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

		public void MouseOut([In] object button, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseOut");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, target);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			NetOffice.OWC10Api._Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.OWC10Api._Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MouseOver([In] object button, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseOver");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, target);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			NetOffice.OWC10Api._Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.OWC10Api._Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newTarget;
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

		public void SelectionChanging([In, MarshalAs(UnmanagedType.IDispatch)] object range)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChanging");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(range);
				return;
			}

			NetOffice.OWC10Api._Range newRange = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, range) as NetOffice.OWC10Api._Range;
			object[] paramsArray = new object[1];
			paramsArray[0] = newRange;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			NetOffice.OWC10Api.Worksheet newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as NetOffice.OWC10Api.Worksheet;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetCalculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			NetOffice.OWC10Api.Worksheet newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as NetOffice.OWC10Api.Worksheet;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			NetOffice.OWC10Api.Worksheet newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as NetOffice.OWC10Api.Worksheet;
			NetOffice.OWC10Api._Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.OWC10Api._Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			NetOffice.OWC10Api.Worksheet newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as NetOffice.OWC10Api.Worksheet;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetFollowHyperlink");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			NetOffice.OWC10Api.Worksheet newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as NetOffice.OWC10Api.Worksheet;
			NetOffice.OWC10Api.Hyperlink newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.OWC10Api.Hyperlink;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StartEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection, initialValue, cancel, errorDescription);
				return;
			}

			object newSelection = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, selection) as object;
			NetOffice.OWC10Api.ByRef newInitialValue = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, initialValue) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newErrorDescription = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, errorDescription) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[4];
			paramsArray[0] = newSelection;
			paramsArray[1] = newInitialValue;
			paramsArray[2] = newCancel;
			paramsArray[3] = newErrorDescription;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.OWC10Api._Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.OWC10Api._Range;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}