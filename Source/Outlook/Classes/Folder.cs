using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Folder_BeforeFolderMoveEventHandler(NetOffice.OutlookApi.MAPIFolder moveTo, ref bool cancel);
	public delegate void Folder_BeforeItemMoveEventHandler(ICOMObject item, NetOffice.OutlookApi.MAPIFolder moveTo, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Folder 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863890.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.MAPIFolderEvents_12))]
	[TypeId("000610F7-0000-0000-C000-000000000046")]
    
    public interface Folder : MAPIFolder, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868895.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event Folder_BeforeFolderMoveEventHandler BeforeFolderMoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869445.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Folder_BeforeItemMoveEventHandler BeforeItemMoveEvent;

        #endregion
    }
}
