using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630B2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ReminderCollectionEvents
	{
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64147)]
		void BeforeReminderShow([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64148)]
		void ReminderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64149)]
		void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64150)]
		void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64151)]
		void ReminderRemove();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64152)]
		void Snooze([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ReminderCollectionEvents_SinkHelper : SinkHelper, ReminderCollectionEvents
	{
		#region Static
		
		public static readonly string Id = "000630B2-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public ReminderCollectionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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
			_eventBinding.RaiseCustomEvent("BeforeReminderShow", ref paramsArray);

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

			NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			_eventBinding.RaiseCustomEvent("ReminderAdd", ref paramsArray);
		}

		public void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReminderChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			_eventBinding.RaiseCustomEvent("ReminderChange", ref paramsArray);
		}

		public void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReminderFire");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			_eventBinding.RaiseCustomEvent("ReminderFire", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("ReminderRemove", ref paramsArray);
		}

		public void Snooze([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Snooze");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reminderObject);
				return;
			}

			NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateObjectFromComProxy(_eventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			_eventBinding.RaiseCustomEvent("Snooze", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}