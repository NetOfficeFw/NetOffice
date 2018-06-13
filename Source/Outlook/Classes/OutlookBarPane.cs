using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OutlookBarPane_BeforeNavigateEventHandler(NetOffice.OutlookApi.OutlookBarShortcut shortcut, ref bool cancel);
	public delegate void OutlookBarPane_BeforeGroupSwitchEventHandler(NetOffice.OutlookApi.OutlookBarGroup toGroup, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OutlookBarPane 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870061.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OutlookBarPaneEvents))]
	[TypeId("00063055-0000-0000-C000-000000000046")]
    public interface OutlookBarPane : _OutlookBarPane, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869977.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarPane_BeforeNavigateEventHandler BeforeNavigateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarPane_BeforeGroupSwitchEventHandler BeforeGroupSwitchEvent;

        #endregion
    }
}
