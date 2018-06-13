using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkListBox_ClickEventHandler();
	public delegate void OlkListBox_DoubleClickEventHandler();
	public delegate void OlkListBox_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkListBox_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkListBox_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkListBox_EnterEventHandler();
	public delegate void OlkListBox_ExitEventHandler(ref bool cancel);
	public delegate void OlkListBox_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkListBox_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkListBox_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkListBox_ChangeEventHandler();
	public delegate void OlkListBox_AfterUpdateEventHandler();
	public delegate void OlkListBox_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkListBox 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863585.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkListBoxEvents))]
	[TypeId("0006F04E-0000-0000-C000-000000000046")]
    public interface OlkListBox : _OlkListBox, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866067.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866412.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869274.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868747.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870174.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870045.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866452.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868095.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866003.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866774.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868533.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861330.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862397.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkListBox_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
