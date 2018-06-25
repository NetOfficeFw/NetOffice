using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_ItemSendEventHandler(ICOMObject item, ref bool cancel);
	public delegate void Application_NewMailEventHandler();
	public delegate void Application_ReminderEventHandler(ICOMObject item);
	public delegate void Application_OptionsPagesAddEventHandler(NetOffice.OutlookApi.PropertyPages pages);
	public delegate void Application_StartupEventHandler();
	public delegate void Application_QuitEventHandler();
	public delegate void Application_AdvancedSearchCompleteEventHandler(NetOffice.OutlookApi.Search searchObject);
	public delegate void Application_AdvancedSearchStoppedEventHandler(NetOffice.OutlookApi.Search searchObject);
	public delegate void Application_MAPILogonCompleteEventHandler();
	public delegate void Application_NewMailExEventHandler(string entryIDCollection);
	public delegate void Application_AttachmentContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.AttachmentSelection attachments);
	public delegate void Application_FolderContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.Folder folder);
	public delegate void Application_StoreContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.Store store);
	public delegate void Application_ShortcutContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.OutlookBarShortcut shortcut);
	public delegate void Application_ViewContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.View view);
	public delegate void Application_ItemContextMenuDisplayEventHandler(NetOffice.OfficeApi.CommandBar commandBar, NetOffice.OutlookApi.Selection selection);
	public delegate void Application_ContextMenuCloseEventHandler(NetOffice.OutlookApi.Enums.OlContextMenu contextMenu);
	public delegate void Application_ItemLoadEventHandler(ICOMObject item);
	public delegate void Application_BeforeFolderSharingDialogEventHandler(NetOffice.OutlookApi.MAPIFolder folderToShare, ref bool cancel);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.Outlook.Behind.Application
    /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866895.aspx </remarks>
    [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.OutlookApi.Behind.Application
    {
        private string _defaultProgId = "Outlook.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Outlook
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Outlook based on given id.
        /// This can be used to target a specific version of Microsoft Outlook.
        /// Example usage:
        /// "Microsoft.Outlook.12" to target Outlook 2007
        /// "Microsoft.Outlook.14" to target Outlook 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Outlook
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Outlook
        /// </summary>
        /// <param name="mode">indicates where is the call coming from</param>
        public ApplicationClass(NetOffice.Callers.InteropCompatibilityClassCreateMode mode)
        {
            if (mode == NetOffice.Callers.InteropCompatibilityClassCreateMode.Direct)
            {
                ICOMObjectInitialize init = (ICOMObjectInitialize)this;
                init.InitializeCOMObject(_defaultProgId);
            }
        }
    }

    /// <summary>
    /// CoClass Application
    /// SupportByVersion Outlook, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866895.aspx </remarks>
    [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Outlook.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(EventContracts.ApplicationEvents), typeof(EventContracts.ApplicationEvents_10), typeof(EventContracts.ApplicationEvents_11))]
	[TypeId("0006F03A-0000-0000-C000-000000000046")]
    public interface Application : _Application, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
    {
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865076.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Application_ItemSendEventHandler ItemSendEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869202.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Application_NewMailEventHandler NewMailEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870058.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Application_ReminderEventHandler ReminderEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868446.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Application_OptionsPagesAddEventHandler OptionsPagesAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869298.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Application_StartupEventHandler StartupEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869760.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Application_QuitEventHandler QuitEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864775.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Application_AdvancedSearchCompleteEventHandler AdvancedSearchCompleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868266.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Application_AdvancedSearchStoppedEventHandler AdvancedSearchStoppedEvent;

		/// <summary>
		/// SupportByVersion Outlook 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869443.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		event Application_MAPILogonCompleteEventHandler MAPILogonCompleteEvent;

		/// <summary>
		/// SupportByVersion Outlook 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863686.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		event Application_NewMailExEventHandler NewMailExEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_AttachmentContextMenuDisplayEventHandler AttachmentContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_FolderContextMenuDisplayEventHandler FolderContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_StoreContextMenuDisplayEventHandler StoreContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_ShortcutContextMenuDisplayEventHandler ShortcutContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_ViewContextMenuDisplayEventHandler ViewContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_ItemContextMenuDisplayEventHandler ItemContextMenuDisplayEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_ContextMenuCloseEventHandler ContextMenuCloseEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868544.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_ItemLoadEventHandler ItemLoadEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869543.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event Application_BeforeFolderSharingDialogEventHandler BeforeFolderSharingDialogEvent;

        #endregion
    }
}
