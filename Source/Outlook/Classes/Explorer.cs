using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Explorer_ActivateEventHandler();
	public delegate void Explorer_FolderSwitchEventHandler();
	public delegate void Explorer_BeforeFolderSwitchEventHandler(ICOMObject newFolder, ref bool cancel);
	public delegate void Explorer_ViewSwitchEventHandler();
	public delegate void Explorer_BeforeViewSwitchEventHandler(object newView, ref bool cancel);
	public delegate void Explorer_DeactivateEventHandler();
	public delegate void Explorer_SelectionChangeEventHandler();
	public delegate void Explorer_CloseEventHandler();
	public delegate void Explorer_BeforeMaximizeEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeMinimizeEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeMoveEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeSizeEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeItemCopyEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeItemCutEventHandler(ref bool cancel);
	public delegate void Explorer_BeforeItemPasteEventHandler(ref object clipboardContent, NetOffice.OutlookApi.MAPIFolder target, ref bool cancel);
	public delegate void Explorer_AttachmentSelectionChangeEventHandler();
	public delegate void Explorer_InlineResponseEventHandler(ICOMObject item);
    public delegate void Explorer_InlineResponseCloseEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Explorer 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860356.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ExplorerEvents), typeof(EventContracts.ExplorerEvents_10))]
	[TypeId("00063050-0000-0000-C000-000000000046")]
    public interface Explorer : _Explorer, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867298.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865625.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_FolderSwitchEventHandler FolderSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868537.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_BeforeFolderSwitchEventHandler BeforeFolderSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868484.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_ViewSwitchEventHandler ViewSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865397.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_BeforeViewSwitchEventHandler BeforeViewSwitchEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866945.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_DeactivateEventHandler DeactivateEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869813.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_SelectionChangeEventHandler SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862184.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorer_CloseEventHandler CloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864743.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeMaximizeEventHandler BeforeMaximizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868043.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeMinimizeEventHandler BeforeMinimizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868815.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeMoveEventHandler BeforeMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862995.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeSizeEventHandler BeforeSizeEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860454.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeItemCopyEventHandler BeforeItemCopyEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867174.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeItemCutEventHandler BeforeItemCutEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868366.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Explorer_BeforeItemPasteEventHandler BeforeItemPasteEvent;

		/// <summary>
		/// SupportByVersion Outlook 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867876.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		event Explorer_AttachmentSelectionChangeEventHandler AttachmentSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229061.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		event Explorer_InlineResponseEventHandler InlineResponseEvent;

        /// <summary>
        /// SupportByVersion Outlook 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229061.aspx </remarks>
        [SupportByVersion("Outlook", 15, 16)]
        event Explorer_InlineResponseCloseEventHandler InlineResponseCloseEvent;

        #endregion
    }
}
