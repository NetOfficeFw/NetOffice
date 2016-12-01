using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.OWC10Api
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("OWC10", 1)]
	[ComImport, Guid("F5B39A87-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IPivotControlEvents
	{
		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6003)]
		void SelectionChange();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6004)]
		void ViewChange([In] object reason);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6007)]
		void DataChange([In] object reason);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6021)]
		void PivotTableChange([In] object reason);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6043)]
		void BeforeQuery();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6044)]
		void Query();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6029)]
		void OnConnect();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6030)]
		void OnDisconnect();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6034)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6032)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6033)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6035)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6005)]
		void Click();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6006)]
		void DblClick();

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1000)]
		void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1001)]
		void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void CommandTipText([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1004)]
		void CommandExecute([In] object command, [In] object succeeded);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1009)]
		void KeyDown([In] object keyCode, [In] object shift);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void KeyUp([In] object keyCode, [In] object shift);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1010)]
		void KeyPress([In] object keyAscii);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1006)]
		void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1005)]
		void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1007)]
		void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1011)]
		void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6045)]
		void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object activeObject, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object arrowMode, [In, MarshalAs(UnmanagedType.IDispatch)] object caretPosition, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByVersionAttribute("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6046)]
		void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription);

		[SupportByVersionAttribute("OWC10", 1)]
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
			_eventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("DataChange", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("PivotTableChange", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("BeforeQuery", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("Query", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("OnConnect", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("OnDisconnect", ref paramsArray);
		}

		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newShift = Convert.ToInt32(shift);
			Int32 newx = Convert.ToInt32(x);
			Int32 newy = Convert.ToInt32(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			_eventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

		public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newShift = Convert.ToInt32(shift);
			Int32 newx = Convert.ToInt32(x);
			Int32 newy = Convert.ToInt32(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			_eventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

		public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newShift = Convert.ToInt32(shift);
			Int32 newx = Convert.ToInt32(x);
			Int32 newy = Convert.ToInt32(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		public void MouseWheel([In] object page, [In] object count)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseWheel");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page, count);
				return;
			}

			bool newPage = Convert.ToBoolean(page);
			Int32 newCount = Convert.ToInt32(count);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			_eventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
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

		public void DblClick()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DblClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
		}

		public void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandEnabled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, enabled);
				return;
			}

			object newCommand = (object)command;
			NetOffice.OWC10Api.ByRef newEnabled = Factory.CreateObjectFromComProxy(_eventClass, enabled) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
			_eventBinding.RaiseCustomEvent("CommandEnabled", ref paramsArray);
		}

		public void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandChecked");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, _checked);
				return;
			}

			object newCommand = (object)command;
			NetOffice.OWC10Api.ByRef newChecked = Factory.CreateObjectFromComProxy(_eventClass, _checked) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			_eventBinding.RaiseCustomEvent("CommandChecked", ref paramsArray);
		}

		public void CommandTipText([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandTipText");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, caption);
				return;
			}

			object newCommand = (object)command;
			NetOffice.OWC10Api.ByRef newCaption = Factory.CreateObjectFromComProxy(_eventClass, caption) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCaption;
			_eventBinding.RaiseCustomEvent("CommandTipText", ref paramsArray);
		}

		public void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandBeforeExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, cancel);
				return;
			}

			object newCommand = (object)command;
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			_eventBinding.RaiseCustomEvent("CommandBeforeExecute", ref paramsArray);
		}

		public void CommandExecute([In] object command, [In] object succeeded)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, succeeded);
				return;
			}

			object newCommand = (object)command;
			bool newSucceeded = Convert.ToBoolean(succeeded);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newSucceeded;
			_eventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
		}

		public void KeyDown([In] object keyCode, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newShift = Convert.ToInt32(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			_eventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);
		}

		public void KeyUp([In] object keyCode, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newShift = Convert.ToInt32(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			_eventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
		}

		public void KeyPress([In] object keyAscii)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii);
				return;
			}

			Int32 newKeyAscii = Convert.ToInt32(keyAscii);
			object[] paramsArray = new object[1];
			paramsArray[0] = newKeyAscii;
			_eventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);
		}

		public void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeKeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift, cancel);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newShift = Convert.ToInt32(shift);
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			paramsArray[2] = newCancel;
			_eventBinding.RaiseCustomEvent("BeforeKeyDown", ref paramsArray);
		}

		public void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeKeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift, cancel);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newShift = Convert.ToInt32(shift);
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			paramsArray[2] = newCancel;
			_eventBinding.RaiseCustomEvent("BeforeKeyUp", ref paramsArray);
		}

		public void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeKeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii, cancel);
				return;
			}

			Int32 newKeyAscii = Convert.ToInt32(keyAscii);
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray[1] = newCancel;
			_eventBinding.RaiseCustomEvent("BeforeKeyPress", ref paramsArray);
		}

		public void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeContextMenu");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(x, y, menu, cancel);
				return;
			}

			Int32 newx = Convert.ToInt32(x);
			Int32 newy = Convert.ToInt32(y);
			NetOffice.OWC10Api.ByRef newMenu = Factory.CreateObjectFromComProxy(_eventClass, menu) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[4];
			paramsArray[0] = newx;
			paramsArray[1] = newy;
			paramsArray[2] = newMenu;
			paramsArray[3] = newCancel;
			_eventBinding.RaiseCustomEvent("BeforeContextMenu", ref paramsArray);
		}

		public void StartEdit([In, MarshalAs(UnmanagedType.IDispatch)] object selection, [In, MarshalAs(UnmanagedType.IDispatch)] object activeObject, [In, MarshalAs(UnmanagedType.IDispatch)] object initialValue, [In, MarshalAs(UnmanagedType.IDispatch)] object arrowMode, [In, MarshalAs(UnmanagedType.IDispatch)] object caretPosition, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StartEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection, activeObject, initialValue, arrowMode, caretPosition, cancel, errorDescription);
				return;
			}

			object newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as object;
			object newActiveObject = Factory.CreateObjectFromComProxy(_eventClass, activeObject) as object;
			NetOffice.OWC10Api.ByRef newInitialValue = Factory.CreateObjectFromComProxy(_eventClass, initialValue) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newArrowMode = Factory.CreateObjectFromComProxy(_eventClass, arrowMode) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCaretPosition = Factory.CreateObjectFromComProxy(_eventClass, caretPosition) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newErrorDescription = Factory.CreateObjectFromComProxy(_eventClass, errorDescription) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[7];
			paramsArray[0] = newSelection;
			paramsArray[1] = newActiveObject;
			paramsArray[2] = newInitialValue;
			paramsArray[3] = newArrowMode;
			paramsArray[4] = newCaretPosition;
			paramsArray[5] = newCancel;
			paramsArray[6] = newErrorDescription;
			_eventBinding.RaiseCustomEvent("StartEdit", ref paramsArray);
		}

		public void EndEdit([In] object accept, [In, MarshalAs(UnmanagedType.IDispatch)] object finalValue, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object errorDescription)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EndEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(accept, finalValue, cancel, errorDescription);
				return;
			}

			bool newAccept = Convert.ToBoolean(accept);
			NetOffice.OWC10Api.ByRef newFinalValue = Factory.CreateObjectFromComProxy(_eventClass, finalValue) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.OWC10Api.ByRef;
			NetOffice.OWC10Api.ByRef newErrorDescription = Factory.CreateObjectFromComProxy(_eventClass, errorDescription) as NetOffice.OWC10Api.ByRef;
			object[] paramsArray = new object[4];
			paramsArray[0] = newAccept;
			paramsArray[1] = newFinalValue;
			paramsArray[2] = newCancel;
			paramsArray[3] = newErrorDescription;
			_eventBinding.RaiseCustomEvent("EndEdit", ref paramsArray);
		}

		public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeScreenTip");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(screenTipText, sourceObject);
				return;
			}

			NetOffice.OWC10Api.ByRef newScreenTipText = Factory.CreateObjectFromComProxy(_eventClass, screenTipText) as NetOffice.OWC10Api.ByRef;
			object newSourceObject = Factory.CreateObjectFromComProxy(_eventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			_eventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}