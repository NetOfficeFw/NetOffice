using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Folders_FolderAddEventHandler(NetOffice.OutlookApi.MAPIFolder folder);
	public delegate void Folders_FolderChangeEventHandler(NetOffice.OutlookApi.MAPIFolder folder);
	public delegate void Folders_FolderRemoveEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Folders 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860950.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.FoldersEvents))]
	[TypeId("00063051-0000-0000-C000-000000000046")]
    public interface Folders : _Folders, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869354.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Folders_FolderAddEventHandler FolderAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869140.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Folders_FolderChangeEventHandler FolderChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867661.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Folders_FolderRemoveEventHandler FolderRemoveEvent;

        #endregion
    }
}
