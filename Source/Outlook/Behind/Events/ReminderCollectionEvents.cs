using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ReminderCollectionEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ReminderCollectionEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ReminderCollectionEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from ReminderCollectionEvents
        /// </summary>
        public static readonly string Id = "000630B2-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public ReminderCollectionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ReminderCollectionEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reminderObject"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reminderObject"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reminderObject"></param>
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

        /// <summary>
        /// 
        /// </summary>
		public void ReminderRemove()
        {
            if (!Validate("ReminderRemove"))
            {     
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ReminderRemove", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reminderObject"></param>
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
}
