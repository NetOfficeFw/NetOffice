using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OutlookBarShortcuts_ShortcutAddEventHandler(NetOffice.OutlookApi.OutlookBarShortcut newShortcut);
	public delegate void OutlookBarShortcuts_BeforeShortcutAddEventHandler(ref bool cancel);
	public delegate void OutlookBarShortcuts_BeforeShortcutRemoveEventHandler(NetOffice.OutlookApi.OutlookBarShortcut shortcut, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OutlookBarShortcuts 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865646.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OutlookBarShortcutsEvents))]
	[TypeId("00063057-0000-0000-C000-000000000046")]
    public interface OutlookBarShortcuts : _OutlookBarShortcuts, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869326.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarShortcuts_ShortcutAddEventHandler ShortcutAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868634.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarShortcuts_BeforeShortcutAddEventHandler BeforeShortcutAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864469.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarShortcuts_BeforeShortcutRemoveEventHandler BeforeShortcutRemoveEvent;

        #endregion
    }
}
