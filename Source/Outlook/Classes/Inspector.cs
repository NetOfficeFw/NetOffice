using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Inspector_ActivateEventHandler();
	public delegate void Inspector_DeactivateEventHandler();
	public delegate void Inspector_CloseEventHandler();
	public delegate void Inspector_BeforeMaximizeEventHandler(ref bool cancel);
	public delegate void Inspector_BeforeMinimizeEventHandler(ref bool cancel);
	public delegate void Inspector_BeforeMoveEventHandler(ref bool cancel);
	public delegate void Inspector_BeforeSizeEventHandler(ref bool cancel);
	public delegate void Inspector_PageChangeEventHandler(ref string activePageName);
	public delegate void Inspector_AttachmentSelectionChangeEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Inspector 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869356.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.InspectorEvents), typeof(EventContracts.InspectorEvents_10))]
	[TypeId("00063058-0000-0000-C000-000000000046")]
    public interface Inspector : _Inspector, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865363.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Inspector_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862214.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Inspector_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865374.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Inspector_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867903.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Inspector_BeforeMaximizeEventHandler BeforeMaximizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868289.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Inspector_BeforeMinimizeEventHandler BeforeMinimizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865042.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Inspector_BeforeMoveEventHandler BeforeMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869786.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Inspector_BeforeSizeEventHandler BeforeSizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869845.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Inspector_PageChangeEventHandler PageChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861296.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event Inspector_AttachmentSelectionChangeEventHandler AttachmentSelectionChangeEvent;

        #endregion
    }
}
