using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("OL10","OL11","OL12","OL14")]
	[ComImport, Guid("000630B2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ReminderCollectionEvents
	{
		[SupportByLibrary("OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64147)]
		void BeforeReminderShow([In] [Out] ref object cancel);

		[SupportByLibrary("OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64148)]
		void ReminderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByLibrary("OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64149)]
		void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByLibrary("OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64150)]
		void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByLibrary("OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64151)]
		void ReminderRemove();

		[SupportByLibrary("OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64152)]
		void Snooze([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ReminderCollectionEvents_SinkHelper : SinkHelper, ReminderCollectionEvents
	{
		#region Static
		
		public static readonly string Id = "000630B2-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ReminderCollectionEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ReminderCollectionEvents Members
		
		public void BeforeReminderShow([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeReminderShow");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[0];
		}

		public void ReminderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReminderAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			LateBindingApi.OutlookApi._Reminder newReminderObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as LateBindingApi.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReminderChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			LateBindingApi.OutlookApi._Reminder newReminderObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as LateBindingApi.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReminderFire");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			LateBindingApi.OutlookApi._Reminder newReminderObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as LateBindingApi.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ReminderRemove()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReminderRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Snooze([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Snooze");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			LateBindingApi.OutlookApi._Reminder newReminderObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as LateBindingApi.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}