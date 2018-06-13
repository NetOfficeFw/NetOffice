using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ReminderCollectionEvents
    /// </summary>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630B2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ReminderCollectionEvents
	{
        /// <summary>
        /// BeforeReminderShow
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64147)]
		void BeforeReminderShow([In] [Out] ref object cancel);

        /// <summary>
        /// ReminderAdd
        /// </summary>
        /// <param name="reminderObject"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64148)]
		void ReminderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

        /// <summary>
        /// ReminderChange
        /// </summary>
        /// <param name="reminderObject"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64149)]
		void ReminderChange([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

        /// <summary>
        /// ReminderFire
        /// </summary>
        /// <param name="reminderObject"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64150)]
		void ReminderFire([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);

        /// <summary>
        /// ReminderRemove
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64151)]
		void ReminderRemove();

        /// <summary>
        /// Snooze
        /// </summary>
        /// <param name="reminderObject"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("reminderObject", typeof(NetOffice.OutlookApi._Reminder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64152)]
		void Snooze([In, MarshalAs(UnmanagedType.IDispatch)] object reminderObject);
	}
}
