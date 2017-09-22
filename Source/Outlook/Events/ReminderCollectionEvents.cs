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
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64147)]
		void BeforeReminderShow([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64148)]
		void ReminderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64149)]
		void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64150)]
		void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64151)]
		void ReminderRemove();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
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
	
		#region Ctor

		public ReminderCollectionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ReminderCollectionEvents
		
		public void BeforeReminderShow([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeReminderShow"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeReminderShow", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		public void ReminderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
        {
            if (!Validate("ReminderAdd"))
            {
                Invoker.ReleaseParamsArray(reminderObject);
                return;
            }

            NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			EventBinding.RaiseCustomEvent("ReminderAdd", ref paramsArray);
		}

		public void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
        {
            if (!Validate("ReminderChange"))
            {
                Invoker.ReleaseParamsArray(reminderObject);
                return;
            }

            NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			EventBinding.RaiseCustomEvent("ReminderChange", ref paramsArray);
		}

		public void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
        {
            if (!Validate("ReminderFire"))
            {
                Invoker.ReleaseParamsArray(reminderObject);
                return;
            }

            NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			EventBinding.RaiseCustomEvent("ReminderFire", ref paramsArray);
		}

		public void ReminderRemove()
        {
            if (!Validate("ReminderRemove"))
            {     
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ReminderRemove", ref paramsArray);
		}

		public void Snooze([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject)
		{
            if (!Validate("Snooze"))
            {
                Invoker.ReleaseParamsArray(reminderObject);
                return;
            }

            NetOffice.OutlookApi._Reminder newReminderObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, reminderObject) as NetOffice.OutlookApi._Reminder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newReminderObject;
			EventBinding.RaiseCustomEvent("Snooze", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}