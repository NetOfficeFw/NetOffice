using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Stores_BeforeStoreRemoveEventHandler(NetOffice.OutlookApi._Store store, ref bool cancel);
	public delegate void Stores_StoreAddEventHandler(NetOffice.OutlookApi._Store store);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Stores 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867405.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.StoresEvents_12))]
	[TypeId("000610C6-0000-0000-C000-000000000046")]
    public interface Stores : _Stores, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868606.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event Stores_BeforeStoreRemoveEventHandler BeforeStoreRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862524.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Stores_StoreAddEventHandler StoreAddEvent;

        #endregion
    }
}
