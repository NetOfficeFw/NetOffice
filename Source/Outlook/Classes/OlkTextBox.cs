using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkTextBox_ClickEventHandler();
	public delegate void OlkTextBox_DoubleClickEventHandler();
	public delegate void OlkTextBox_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTextBox_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTextBox_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkTextBox_EnterEventHandler();
	public delegate void OlkTextBox_ExitEventHandler(ref bool cancel);
	public delegate void OlkTextBox_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkTextBox_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkTextBox_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkTextBox_ChangeEventHandler();
	public delegate void OlkTextBox_AfterUpdateEventHandler();
	public delegate void OlkTextBox_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkTextBox 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867552.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkTextBoxEvents))]
	[TypeId("0006F068-0000-0000-C000-000000000046")]
    public interface OlkTextBox : _OlkTextBox, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868548.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861869.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868623.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864226.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866264.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869484.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869710.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868368.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863915.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866420.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869076.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869973.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868860.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkTextBox_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
