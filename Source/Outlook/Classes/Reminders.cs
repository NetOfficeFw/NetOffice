using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Reminders_BeforeReminderShowEventHandler(ref bool cancel);
	public delegate void Reminders_ReminderAddEventHandler(NetOffice.OutlookApi._Reminder reminderObject);
	public delegate void Reminders_ReminderChangeEventHandler(NetOffice.OutlookApi._Reminder reminderObject);
	public delegate void Reminders_ReminderFireEventHandler(NetOffice.OutlookApi._Reminder reminderObject);
	public delegate void Reminders_ReminderRemoveEventHandler();
	public delegate void Reminders_SnoozeEventHandler(NetOffice.OutlookApi._Reminder reminderObject);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Reminders 
	/// SupportByVersion Outlook, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866017.aspx </remarks>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ReminderCollectionEvents))]
	[TypeId("0006F029-0000-0000-C000-000000000046")]
    public interface Reminders : _Reminders, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867326.aspx </remarks>
        [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Reminders_BeforeReminderShowEventHandler BeforeReminderShowEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869105.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Reminders_ReminderAddEventHandler ReminderAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863669.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Reminders_ReminderChangeEventHandler ReminderChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866477.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Reminders_ReminderFireEventHandler ReminderFireEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869874.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Reminders_ReminderRemoveEventHandler ReminderRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862485.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Reminders_SnoozeEventHandler SnoozeEvent;

        #endregion
    }
}
