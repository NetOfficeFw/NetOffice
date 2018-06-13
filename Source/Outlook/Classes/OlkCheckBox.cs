using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkCheckBox_ClickEventHandler();
	public delegate void OlkCheckBox_DoubleClickEventHandler();
	public delegate void OlkCheckBox_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCheckBox_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCheckBox_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCheckBox_EnterEventHandler();
	public delegate void OlkCheckBox_ExitEventHandler(ref bool cancel);
	public delegate void OlkCheckBox_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkCheckBox_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkCheckBox_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkCheckBox_ChangeEventHandler();
	public delegate void OlkCheckBox_AfterUpdateEventHandler();
	public delegate void OlkCheckBox_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkCheckBox 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866887.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkCheckBoxEvents))]
	[TypeId("0006F04C-0000-0000-C000-000000000046")]
    public interface OlkCheckBox : _OlkCheckBox, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866392.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868085.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861910.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861031.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860945.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862996.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868405.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869384.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868465.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864423.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865074.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868275.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869554.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCheckBox_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
