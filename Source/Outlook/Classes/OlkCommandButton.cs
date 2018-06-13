using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkCommandButton_ClickEventHandler();
	public delegate void OlkCommandButton_DoubleClickEventHandler();
	public delegate void OlkCommandButton_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCommandButton_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCommandButton_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCommandButton_EnterEventHandler();
	public delegate void OlkCommandButton_ExitEventHandler(ref bool cancel);
	public delegate void OlkCommandButton_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkCommandButton_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkCommandButton_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkCommandButton_AfterUpdateEventHandler();
	public delegate void OlkCommandButton_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkCommandButton 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868781.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkCommandButtonEvents))]
	[TypeId("0006F04A-0000-0000-C000-000000000046")]
    public interface OlkCommandButton : _OlkCommandButton, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863427.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869593.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868329.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862982.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860671.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868553.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868836.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865822.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869029.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865850.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863036.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865606.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCommandButton_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
