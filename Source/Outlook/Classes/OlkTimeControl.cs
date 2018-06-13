using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkTimeControl_ClickEventHandler();
	public delegate void OlkTimeControl_DoubleClickEventHandler();
	public delegate void OlkTimeControl_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTimeControl_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTimeControl_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTimeControl_EnterEventHandler();
	public delegate void OlkTimeControl_ExitEventHandler(ref bool cancel);
	public delegate void OlkTimeControl_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkTimeControl_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkTimeControl_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkTimeControl_ChangeEventHandler();
	public delegate void OlkTimeControl_DropButtonClickEventHandler();
	public delegate void OlkTimeControl_AfterUpdateEventHandler();
	public delegate void OlkTimeControl_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkTimeControl 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868612.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkTimeControlEvents))]
	[TypeId("0006F051-0000-0000-C000-000000000046")]
    public interface OlkTimeControl : _OlkTimeControl, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866709.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869446.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865862.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865094.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870088.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862681.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860380.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861291.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865313.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868629.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867578.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862791.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_DropButtonClickEventHandler DropButtonClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865068.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868825.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTimeControl_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
