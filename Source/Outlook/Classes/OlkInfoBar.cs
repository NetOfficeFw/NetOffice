using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkInfoBar_ClickEventHandler();
	public delegate void OlkInfoBar_DoubleClickEventHandler();
	public delegate void OlkInfoBar_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkInfoBar_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkInfoBar_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkInfoBar 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861894.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkInfoBarEvents))]
	[TypeId("0006F054-0000-0000-C000-000000000046")]
    public interface OlkInfoBar : _OlkInfoBar, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860621.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkInfoBar_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861240.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkInfoBar_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868263.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkInfoBar_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868396.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkInfoBar_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869434.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkInfoBar_MouseUpEventHandler MouseUpEvent;

        #endregion
    }
}
