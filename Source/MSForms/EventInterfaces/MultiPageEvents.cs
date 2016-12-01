using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.MSFormsApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("MSForms", 2)]
	[ComImport, Guid("7B020EC8-AF6C-11CE-9F46-00AA00574A4F"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MultiPageEvents
	{
		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(768)]
		void AddControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object state, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Change();

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click([In] object index);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-608)]
		void Error([In] object index, [In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(770)]
		void Layout([In] object index);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(771)]
		void RemoveControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(772)]
		void Scroll([In] object index, [In] object actionX, [In] object actionY, [In] object requestDx, [In] object requestDy, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDx, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDy);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(773)]
		void Zoom([In] object index, [In] [Out] ref object percent);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MultiPageEvents_SinkHelper : SinkHelper, MultiPageEvents
	{
		#region Static
		
		public static readonly string Id = "7B020EC8-AF6C-11CE-9F46-00AA00574A4F";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public MultiPageEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region MultiPageEvents Members
		
		public void AddControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AddControl");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, control);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.Control newControl = Factory.CreateObjectFromComProxy(_eventClass, control) as NetOffice.MSFormsApi.Control;
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray[1] = newControl;
			_eventBinding.RaiseCustomEvent("AddControl", ref paramsArray);
		}

		public void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object state, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDragOver");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, cancel, control, data, x, y, state, effect, shift);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			NetOffice.MSFormsApi.Control newControl = Factory.CreateObjectFromComProxy(_eventClass, control) as NetOffice.MSFormsApi.Control;
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateObjectFromComProxy(_eventClass, data) as NetOffice.MSFormsApi.DataObject;
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			NetOffice.MSFormsApi.Enums.fmDragState newState = (NetOffice.MSFormsApi.Enums.fmDragState)state;
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateObjectFromComProxy(_eventClass, effect) as NetOffice.MSFormsApi.ReturnEffect;
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[9];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			paramsArray[2] = newControl;
			paramsArray[3] = newData;
			paramsArray[4] = newX;
			paramsArray[5] = newY;
			paramsArray[6] = newState;
			paramsArray[7] = newEffect;
			paramsArray[8] = newShift;
			_eventBinding.RaiseCustomEvent("BeforeDragOver", ref paramsArray);
		}

		public void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDropOrPaste");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, cancel, control, action, data, x, y, effect, shift);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			NetOffice.MSFormsApi.Control newControl = Factory.CreateObjectFromComProxy(_eventClass, control) as NetOffice.MSFormsApi.Control;
			NetOffice.MSFormsApi.Enums.fmAction newAction = (NetOffice.MSFormsApi.Enums.fmAction)action;
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateObjectFromComProxy(_eventClass, data) as NetOffice.MSFormsApi.DataObject;
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateObjectFromComProxy(_eventClass, effect) as NetOffice.MSFormsApi.ReturnEffect;
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[9];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			paramsArray[2] = newControl;
			paramsArray[3] = newAction;
			paramsArray[4] = newData;
			paramsArray[5] = newX;
			paramsArray[6] = newY;
			paramsArray[7] = newEffect;
			paramsArray[8] = newShift;
			_eventBinding.RaiseCustomEvent("BeforeDropOrPaste", ref paramsArray);
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

		public void Click([In] object index)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Click");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			object[] paramsArray = new object[1];
			paramsArray[0] = newIndex;
			_eventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		public void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DblClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, cancel);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			_eventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
		}

		public void Error([In] object index, [In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Error");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, number, description, sCode, source, helpFile, helpContext, cancelDisplay);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			Int16 newNumber = Convert.ToInt16(number);
			NetOffice.MSFormsApi.ReturnString newDescription = Factory.CreateObjectFromComProxy(_eventClass, description) as NetOffice.MSFormsApi.ReturnString;
			Int32 newSCode = Convert.ToInt32(sCode);
			string newSource = Convert.ToString(source);
			string newHelpFile = Convert.ToString(helpFile);
			Int32 newHelpContext = Convert.ToInt32(helpContext);
			NetOffice.MSFormsApi.ReturnBoolean newCancelDisplay = Factory.CreateObjectFromComProxy(_eventClass, cancelDisplay) as NetOffice.MSFormsApi.ReturnBoolean;
			object[] paramsArray = new object[8];
			paramsArray[0] = newIndex;
			paramsArray[1] = newNumber;
			paramsArray[2] = newDescription;
			paramsArray[3] = newSCode;
			paramsArray[4] = newSource;
			paramsArray[5] = newHelpFile;
			paramsArray[6] = newHelpContext;
			paramsArray[7] = newCancelDisplay;
			_eventBinding.RaiseCustomEvent("Error", ref paramsArray);
		}

		public void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			NetOffice.MSFormsApi.ReturnInteger newKeyCode = Factory.CreateObjectFromComProxy(_eventClass, keyCode) as NetOffice.MSFormsApi.ReturnInteger;
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			_eventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);
		}

		public void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii);
				return;
			}

			NetOffice.MSFormsApi.ReturnInteger newKeyAscii = Factory.CreateObjectFromComProxy(_eventClass, keyAscii) as NetOffice.MSFormsApi.ReturnInteger;
			object[] paramsArray = new object[1];
			paramsArray[0] = newKeyAscii;
			_eventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);
		}

		public void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			NetOffice.MSFormsApi.ReturnInteger newKeyCode = Factory.CreateObjectFromComProxy(_eventClass, keyCode) as NetOffice.MSFormsApi.ReturnInteger;
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			_eventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
		}

		public void Layout([In] object index)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Layout");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			object[] paramsArray = new object[1];
			paramsArray[0] = newIndex;
			_eventBinding.RaiseCustomEvent("Layout", ref paramsArray);
		}

		public void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, button, shift, x, y);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			Int16 newButton = Convert.ToInt16(button);
			Int16 newShift = Convert.ToInt16(shift);
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			_eventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

		public void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, button, shift, x, y);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			Int16 newButton = Convert.ToInt16(button);
			Int16 newShift = Convert.ToInt16(shift);
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			_eventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

		public void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, button, shift, x, y);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			Int16 newButton = Convert.ToInt16(button);
			Int16 newShift = Convert.ToInt16(shift);
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		public void RemoveControl([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object control)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RemoveControl");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, control);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.Control newControl = Factory.CreateObjectFromComProxy(_eventClass, control) as NetOffice.MSFormsApi.Control;
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray[1] = newControl;
			_eventBinding.RaiseCustomEvent("RemoveControl", ref paramsArray);
		}

		public void Scroll([In] object index, [In] object actionX, [In] object actionY, [In] object requestDx, [In] object requestDy, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDx, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDy)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Scroll");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, actionX, actionY, requestDx, requestDy, actualDx, actualDy);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.Enums.fmScrollAction newActionX = (NetOffice.MSFormsApi.Enums.fmScrollAction)actionX;
			NetOffice.MSFormsApi.Enums.fmScrollAction newActionY = (NetOffice.MSFormsApi.Enums.fmScrollAction)actionY;
			Single newRequestDx = Convert.ToSingle(requestDx);
			Single newRequestDy = Convert.ToSingle(requestDy);
			NetOffice.MSFormsApi.ReturnSingle newActualDx = Factory.CreateObjectFromComProxy(_eventClass, actualDx) as NetOffice.MSFormsApi.ReturnSingle;
			NetOffice.MSFormsApi.ReturnSingle newActualDy = Factory.CreateObjectFromComProxy(_eventClass, actualDy) as NetOffice.MSFormsApi.ReturnSingle;
			object[] paramsArray = new object[7];
			paramsArray[0] = newIndex;
			paramsArray[1] = newActionX;
			paramsArray[2] = newActionY;
			paramsArray[3] = newRequestDx;
			paramsArray[4] = newRequestDy;
			paramsArray[5] = newActualDx;
			paramsArray[6] = newActualDy;
			_eventBinding.RaiseCustomEvent("Scroll", ref paramsArray);
		}

		public void Zoom([In] object index, [In] [Out] ref object percent)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Zoom");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(index, percent);
				return;
			}

			Int32 newIndex = Convert.ToInt32(index);
			object[] paramsArray = new object[2];
			paramsArray[0] = newIndex;
			paramsArray.SetValue(percent, 1);
			_eventBinding.RaiseCustomEvent("Zoom", ref paramsArray);

			percent = (Int16)paramsArray[1];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}