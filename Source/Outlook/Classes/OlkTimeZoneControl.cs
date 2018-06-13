using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkTimeZoneControl_ClickEventHandler();
	public delegate void OlkTimeZoneControl_DoubleClickEventHandler();
	public delegate void OlkTimeZoneControl_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTimeZoneControl_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTimeZoneControl_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTimeZoneControl_EnterEventHandler();
	public delegate void OlkTimeZoneControl_ExitEventHandler(ref bool cancel);
	public delegate void OlkTimeZoneControl_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkTimeZoneControl_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkTimeZoneControl_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkTimeZoneControl_ChangeEventHandler();
	public delegate void OlkTimeZoneControl_DropButtonClickEventHandler();
	public delegate void OlkTimeZoneControl_AfterUpdateEventHandler();
	public delegate void OlkTimeZoneControl_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkTimeZoneControl 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862219.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkTimeZoneControlEvents))]
	[TypeId("0006F059-0000-0000-C000-000000000046")]
    public interface OlkTimeZoneControl : _OlkTimeZoneControl, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864773.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862978.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865592.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863912.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867832.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862463.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869423.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861566.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864702.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860637.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863670.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864490.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_DropButtonClickEventHandler DropButtonClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868638.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869906.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeZoneControl_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
