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
	[ComImport, Guid("4C5992A5-6926-101B-9992-00000B65C6F9"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ImageEvents
	{
		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeDragOver([In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object dragState, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforeDropOrPaste([In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-608)]
		void Error([In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ImageEvents_SinkHelper : SinkHelper, ImageEvents
	{
		#region Static
		
		public static readonly string Id = "4C5992A5-6926-101B-9992-00000B65C6F9";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ImageEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region ImageEvents Members
		
		public void BeforeDragOver([In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object dragState, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDragOver");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, data, x, y, dragState, effect, shift);
				return;
			}

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateObjectFromComProxy(_eventClass, data) as NetOffice.MSFormsApi.DataObject;
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			NetOffice.MSFormsApi.Enums.fmDragState newDragState = (NetOffice.MSFormsApi.Enums.fmDragState)dragState;
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateObjectFromComProxy(_eventClass, effect) as NetOffice.MSFormsApi.ReturnEffect;
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[7];
			paramsArray[0] = newCancel;
			paramsArray[1] = newData;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			paramsArray[4] = newDragState;
			paramsArray[5] = newEffect;
			paramsArray[6] = newShift;
			_eventBinding.RaiseCustomEvent("BeforeDragOver", ref paramsArray);
		}

		public void BeforeDropOrPaste([In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDropOrPaste");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, action, data, x, y, effect, shift);
				return;
			}

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			NetOffice.MSFormsApi.Enums.fmAction newAction = (NetOffice.MSFormsApi.Enums.fmAction)action;
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateObjectFromComProxy(_eventClass, data) as NetOffice.MSFormsApi.DataObject;
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateObjectFromComProxy(_eventClass, effect) as NetOffice.MSFormsApi.ReturnEffect;
			Int16 newShift = Convert.ToInt16(shift);
			object[] paramsArray = new object[7];
			paramsArray[0] = newCancel;
			paramsArray[1] = newAction;
			paramsArray[2] = newData;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			paramsArray[5] = newEffect;
			paramsArray[6] = newShift;
			_eventBinding.RaiseCustomEvent("BeforeDropOrPaste", ref paramsArray);
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

		public void DblClick([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DblClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			_eventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
		}

		public void Error([In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Error");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(number, description, sCode, source, helpFile, helpContext, cancelDisplay);
				return;
			}

			Int16 newNumber = Convert.ToInt16(number);
			NetOffice.MSFormsApi.ReturnString newDescription = Factory.CreateObjectFromComProxy(_eventClass, description) as NetOffice.MSFormsApi.ReturnString;
			Int32 newSCode = Convert.ToInt32(sCode);
			string newSource = Convert.ToString(source);
			string newHelpFile = Convert.ToString(helpFile);
			Int32 newHelpContext = Convert.ToInt32(helpContext);
			NetOffice.MSFormsApi.ReturnBoolean newCancelDisplay = Factory.CreateObjectFromComProxy(_eventClass, cancelDisplay) as NetOffice.MSFormsApi.ReturnBoolean;
			object[] paramsArray = new object[7];
			paramsArray[0] = newNumber;
			paramsArray[1] = newDescription;
			paramsArray[2] = newSCode;
			paramsArray[3] = newSource;
			paramsArray[4] = newHelpFile;
			paramsArray[5] = newHelpContext;
			paramsArray[6] = newCancelDisplay;
			_eventBinding.RaiseCustomEvent("Error", ref paramsArray);
		}

		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int16 newButton = Convert.ToInt16(button);
			Int16 newShift = Convert.ToInt16(shift);
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
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

			Int16 newButton = Convert.ToInt16(button);
			Int16 newShift = Convert.ToInt16(shift);
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
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

			Int16 newButton = Convert.ToInt16(button);
			Int16 newShift = Convert.ToInt16(shift);
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}