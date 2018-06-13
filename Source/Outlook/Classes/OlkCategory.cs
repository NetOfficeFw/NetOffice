using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OlkCategory_ClickEventHandler();
	public delegate void OlkCategory_DoubleClickEventHandler();
	public delegate void OlkCategory_MouseDownEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCategory_MouseMoveEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCategory_MouseUpEventHandler(NetOffice.OutlookApi.Enums.OlMouseButton button, NetOffice.OutlookApi.Enums.OlShiftState shift, Single x, Single y);
	public delegate void OlkCategory_EnterEventHandler();
	public delegate void OlkCategory_ExitEventHandler(ref bool cancel);
	public delegate void OlkCategory_KeyDownEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkCategory_KeyPressEventHandler(ref Int32 keyAscii);
	public delegate void OlkCategory_KeyUpEventHandler(ref Int32 keyCode, NetOffice.OutlookApi.Enums.OlShiftState shift);
	public delegate void OlkCategory_ChangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OlkCategory 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869980.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OlkCategoryEvents))]
	[TypeId("0006F053-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.OlkCategory))]
    public interface OlkCategory : _OlkCategory, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866040.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861328.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_DoubleClickEventHandler DoubleClickEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869033.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865335.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868230.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869676.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868801.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869464.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861572.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869167.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867627.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event OlkCategory_ChangeEventHandler ChangeEvent;

        #endregion
    }
}
