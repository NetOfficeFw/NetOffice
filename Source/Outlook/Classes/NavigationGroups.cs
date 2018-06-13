using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void NavigationGroups_SelectedChangeEventHandler(NetOffice.OutlookApi.NavigationFolder navigationFolder);
	public delegate void NavigationGroups_NavigationFolderAddEventHandler(NetOffice.OutlookApi.NavigationFolder navigationFolder);
	public delegate void NavigationGroups_NavigationFolderRemoveEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass NavigationGroups 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860649.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.NavigationGroupsEvents_12))]
	[TypeId("000610F4-0000-0000-C000-000000000046")]
    public interface NavigationGroups : _NavigationGroups, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869729.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event NavigationGroups_SelectedChangeEventHandler SelectedChangeEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868621.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event NavigationGroups_NavigationFolderAddEventHandler NavigationFolderAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862126.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event NavigationGroups_NavigationFolderRemoveEventHandler NavigationFolderRemoveEvent;

        #endregion
    }
}
