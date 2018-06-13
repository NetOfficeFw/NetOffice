using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkBusinessCardControl_ClickEventHandler();
	public delegate void OlkBusinessCardControl_DoubleClickEventHandler();
	public delegate void OlkBusinessCardControl_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkBusinessCardControl_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkBusinessCardControl_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkBusinessCardControl 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868063.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkBusinessCardControlEvents))]
	[TypeId("0006F050-0000-0000-C000-000000000046")]
    public interface OlkBusinessCardControl : _OlkBusinessCardControl, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863387.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkBusinessCardControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867359.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkBusinessCardControl_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862468.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkBusinessCardControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868035.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkBusinessCardControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867371.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkBusinessCardControl_MouseUpEventHandler MouseUpEvent;

        #endregion
    }
}
