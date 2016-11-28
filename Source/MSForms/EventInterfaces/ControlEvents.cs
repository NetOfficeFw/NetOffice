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
	[ComImport, Guid("9A4BBF53-4E46-101B-8BBD-00AA003E3B29"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ControlEvents
	{
		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384830)]
		void Enter();

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384829)]
		void Exit([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384831)]
		void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384832)]
		void AfterUpdate();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ControlEvents_SinkHelper : SinkHelper, ControlEvents
	{
		#region Static
		
		public static readonly string Id = "9A4BBF53-4E46-101B-8BBD-00AA003E3B29";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ControlEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region ControlEvents Members
		
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

		public void Exit([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Exit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			_eventBinding.RaiseCustomEvent("Exit", ref paramsArray);
		}

		public void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as NetOffice.MSFormsApi.ReturnBoolean;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			_eventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);
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

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}