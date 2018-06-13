using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkContactPhoto_ClickEventHandler();
	public delegate void OlkContactPhoto_DoubleClickEventHandler();
	public delegate void OlkContactPhoto_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState Shift, Single x, Single y);
	public delegate void OlkContactPhoto_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkContactPhoto_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkContactPhoto_EnterEventHandler();
	public delegate void OlkContactPhoto_ExitEventHandler(ref bool cancel);
	public delegate void OlkContactPhoto_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkContactPhoto_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkContactPhoto_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkContactPhoto_ChangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkContactPhoto 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869806.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkContactPhotoEvents))]
	[TypeId("0006F04F-0000-0000-C000-000000000046")]
    public interface OlkContactPhoto : _OlkContactPhoto, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864215.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864796.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869332.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869272.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867093.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868566.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867520.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865644.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864241.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869803.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863908.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkContactPhoto_ChangeEventHandler ChangeEvent;

        #endregion
    }
}
