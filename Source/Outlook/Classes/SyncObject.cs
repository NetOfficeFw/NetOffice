using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void SyncObject_SyncStartEventHandler();
	public delegate void SyncObject_ProgressEventHandler(NetOffice.OutlookApi.Enums.OlSyncState state, string description, Int32 value, Int32 max);
	public delegate void SyncObject_OnErrorEventHandler(Int32 code, string description);
	public delegate void SyncObject_SyncEndEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass SyncObject 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860720.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.SyncObjectEvents))]
	[TypeId("00063084-0000-0000-C000-000000000046")]
    public interface SyncObject : _SyncObject, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862356.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event SyncObject_SyncStartEventHandler SyncStartEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865672.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event SyncObject_ProgressEventHandler ProgressEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862157.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event SyncObject_OnErrorEventHandler OnErrorEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866270.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event SyncObject_SyncEndEventHandler SyncEndEvent;

        #endregion
    }
}
