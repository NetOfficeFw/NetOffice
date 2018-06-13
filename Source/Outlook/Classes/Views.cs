using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Views_ViewAddEventHandler(NetOffice.OutlookApi.View view);
	public delegate void Views_ViewRemoveEventHandler(NetOffice.OutlookApi.View view);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Views 
	/// SupportByVersion Outlook, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865619.aspx </remarks>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ViewsEvents))]
	[TypeId("0006F027-0000-0000-C000-000000000046")]
    public interface Views : _Views, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867711.aspx </remarks>
        [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Views_ViewAddEventHandler ViewAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868250.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Views_ViewRemoveEventHandler ViewRemoveEvent;

        #endregion
    }
}
