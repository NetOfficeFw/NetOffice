using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkSenderPhoto_ClickEventHandler();
	public delegate void OlkSenderPhoto_DoubleClickEventHandler();
	public delegate void OlkSenderPhoto_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkSenderPhoto_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkSenderPhoto_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkSenderPhoto_ChangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkSenderPhoto 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860658.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkSenderPhotoEvents))]
	[TypeId("0006F058-0000-0000-C000-000000000046")]
    public interface OlkSenderPhoto : _OlkSenderPhoto, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860412.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkSenderPhoto_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864769.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkSenderPhoto_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867671.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkSenderPhoto_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867570.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkSenderPhoto_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867149.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkSenderPhoto_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868334.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkSenderPhoto_ChangeEventHandler ChangeEvent;

        #endregion
    }
}
