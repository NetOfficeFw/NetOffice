using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkOptionButton_ClickEventHandler();
	public delegate void OlkOptionButton_DoubleClickEventHandler();
	public delegate void OlkOptionButton_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkOptionButton_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkOptionButton_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkOptionButton_EnterEventHandler();
	public delegate void OlkOptionButton_ExitEventHandler(ref bool cancel);
	public delegate void OlkOptionButton_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkOptionButton_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkOptionButton_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkOptionButton_ChangeEventHandler();
	public delegate void OlkOptionButton_AfterUpdateEventHandler();
	public delegate void OlkOptionButton_BeforeUpdateEventHandler(ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkOptionButton 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868387.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkOptionButtonEvents))]
	[TypeId("0006F04B-0000-0000-C000-000000000046")]
    public interface OlkOptionButton : _OlkOptionButton, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870177.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860956.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868365.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863030.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869683.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868415.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862490.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869876.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869168.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868491.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869416.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868453.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868372.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkOptionButton_BeforeUpdateEventHandler BeforeUpdateEvent;

        #endregion
    }
}
