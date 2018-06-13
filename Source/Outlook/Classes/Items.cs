using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Items_ItemAddEventHandler(ICOMObject item);
	public delegate void Items_ItemChangeEventHandler(ICOMObject item);
	public delegate void Items_ItemRemoveEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Items 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863652.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ItemsEvents))]
	[TypeId("00063052-0000-0000-C000-000000000046")]
    public interface Items : _Items, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869609.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Items_ItemAddEventHandler ItemAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865866.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Items_ItemChangeEventHandler ItemChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868911.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Items_ItemRemoveEventHandler ItemRemoveEvent;

        #endregion
    }
}
