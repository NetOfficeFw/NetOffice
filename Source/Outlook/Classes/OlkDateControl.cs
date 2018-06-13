using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkDateControl_ClickEventHandler();
	public delegate void OlkDateControl_DoubleClickEventHandler();
	public delegate void OlkDateControl_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkDateControl_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkDateControl_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkDateControl_EnterEventHandler();
	public delegate void OlkDateControl_ExitEventHandler(ref bool cancel);
	public delegate void OlkDateControl_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkDateControl_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkDateControl_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkDateControl_ChangeEventHandler();
	public delegate void OlkDateControl_DropButtonClickEventHandler();
	public delegate void OlkDateControl_AfterUpdateEventHandler();
	public delegate void OlkDateControl_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkDateControl 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868818.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkDateControlEvents))]
	[TypeId("0006F056-0000-0000-C000-000000000046")]
    public interface OlkDateControl : _OlkDateControl, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869739.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861813.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869518.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868326.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868492.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862109.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866191.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867492.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865350.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866753.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861622.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864198.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_DropButtonClickEventHandler DropButtonClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866407.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862404.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkDateControl_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
