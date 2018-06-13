using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OutlookBarGroups_GroupAddEventHandler(NetOffice.OutlookApi.OutlookBarGroup newGroup);
	public delegate void OutlookBarGroups_BeforeGroupAddEventHandler(ref bool cancel);
	public delegate void OutlookBarGroups_BeforeGroupRemoveEventHandler(NetOffice.OutlookApi.OutlookBarGroup group, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OutlookBarGroups 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868789.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OutlookBarGroupsEvents))]
	[TypeId("00063056-0000-0000-C000-000000000046")]
    public interface OutlookBarGroups : _OutlookBarGroups, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865659.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarGroups_GroupAddEventHandler GroupAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866940.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarGroups_BeforeGroupAddEventHandler BeforeGroupAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868646.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event OutlookBarGroups_BeforeGroupRemoveEventHandler BeforeGroupRemoveEvent;

        #endregion
    }
}
