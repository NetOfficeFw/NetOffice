using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Results_ItemAddEventHandler(ICOMObject item);
	public delegate void Results_ItemChangeEventHandler(ICOMObject item);
	public delegate void Results_ItemRemoveEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Results 
	/// SupportByVersion Outlook, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865331.aspx </remarks>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ResultsEvents))]
	[TypeId("00061039-0000-0000-C000-000000000046")]
    public interface Results : _Results, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868737.aspx </remarks>
        [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Results_ItemAddEventHandler ItemAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861551.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Results_ItemChangeEventHandler ItemChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867866.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Results_ItemRemoveEventHandler ItemRemoveEvent;

        #endregion
    }
}
