using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkComboBox_ClickEventHandler();
	public delegate void OlkComboBox_DoubleClickEventHandler();
	public delegate void OlkComboBox_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkComboBox_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkComboBox_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkComboBox_EnterEventHandler();
	public delegate void OlkComboBox_ExitEventHandler(ref bool cancel);
	public delegate void OlkComboBox_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkComboBox_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkComboBox_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkComboBox_ChangeEventHandler();
	public delegate void OlkComboBox_DropButtonClickEventHandler();
	public delegate void OlkComboBox_AfterUpdateEventHandler();
	public delegate void OlkComboBox_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkComboBox 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867596.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkComboBoxEvents))]
	[TypeId("0006F04D-0000-0000-C000-000000000046")]
    public interface OlkComboBox : _OlkComboBox, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868301.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860953.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866981.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869230.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866284.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866971.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869165.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864242.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868516.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862393.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868856.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868193.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_DropButtonClickEventHandler DropButtonClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869239.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870078.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkComboBox_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
