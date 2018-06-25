using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_WindowActivateEventHandler(NetOffice.PublisherApi.Window wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.PublisherApi.Window wn);
	public delegate void Application_WindowPageChangeEventHandler(NetOffice.PublisherApi.View vw);
	public delegate void Application_QuitEventHandler();
	public delegate void Application_NewDocumentEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_DocumentOpenEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_DocumentBeforeCloseEventHandler(NetOffice.PublisherApi._Document doc, ref bool cancel);
	public delegate void Application_MailMergeAfterMergeEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeAfterRecordMergeEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeBeforeMergeEventHandler(NetOffice.PublisherApi._Document doc, Int32 startRecord, Int32 endRecord, ref bool cancel);
	public delegate void Application_MailMergeBeforeRecordMergeEventHandler(NetOffice.PublisherApi._Document doc, ref bool cancel);
	public delegate void Application_MailMergeDataSourceLoadEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeWizardSendToCustomEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeWizardStateChangeEventHandler(NetOffice.PublisherApi._Document doc, Int32 fromState);
	public delegate void Application_MailMergeDataSourceValidateEventHandler(NetOffice.PublisherApi._Document doc, ref bool handled);
	public delegate void Application_MailMergeInsertBarcodeEventHandler(NetOffice.PublisherApi._Document doc, ref bool okToInsert);
	public delegate void Application_MailMergeRecipientListCloseEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_MailMergeGenerateBarcodeEventHandler(NetOffice.PublisherApi._Document doc, ref string bstrString);
	public delegate void Application_MailMergeWizardFollowUpCustomEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_BeforePrintEventHandler(NetOffice.PublisherApi._Document doc, ref bool cancel);
	public delegate void Application_AfterPrintEventHandler(NetOffice.PublisherApi._Document doc);
	public delegate void Application_ShowCatalogUIEventHandler();
	public delegate void Application_HideCatalogUIEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.PublisherApi.Behind.Application
    /// SupportByVersion Publisher 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("Excel", 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.PublisherApi.Behind.Application
    {
        private string _defaultProgId = "Publisher.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Publisher
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Publisher based on given id.
        /// This can be used to target a specific version of Microsoft Publisher.
        /// Example usage:
        /// "Microsoft.Publisher.12" to target Publisher 2007
        /// "Microsoft.Publisher.14" to target Publisher 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Publisher
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Publisher
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
    /// SupportByVersion Publisher, 14,15,16
    /// </summary>
    [SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Publisher.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(EventContracts.ApplicationEvents))]
	[TypeId("0002123D-0000-0000-C000-000000000046")]
    public interface Application : _Application, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
	{
		#region Events

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_WindowActivateEventHandler WindowActivateEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_WindowDeactivateEventHandler WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_WindowPageChangeEventHandler WindowPageChangeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_QuitEventHandler QuitEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_NewDocumentEventHandler NewDocumentEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_DocumentOpenEventHandler DocumentOpenEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_DocumentBeforeCloseEventHandler DocumentBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeAfterMergeEventHandler MailMergeAfterMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeAfterRecordMergeEventHandler MailMergeAfterRecordMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeBeforeMergeEventHandler MailMergeBeforeMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeBeforeRecordMergeEventHandler MailMergeBeforeRecordMergeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoadEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeWizardSendToCustomEventHandler MailMergeWizardSendToCustomEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeWizardStateChangeEventHandler MailMergeWizardStateChangeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidateEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeInsertBarcodeEventHandler MailMergeInsertBarcodeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeRecipientListCloseEventHandler MailMergeRecipientListCloseEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeGenerateBarcodeEventHandler MailMergeGenerateBarcodeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_MailMergeWizardFollowUpCustomEventHandler MailMergeWizardFollowUpCustomEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_BeforePrintEventHandler BeforePrintEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_AfterPrintEventHandler AfterPrintEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_ShowCatalogUIEventHandler ShowCatalogUIEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Application_HideCatalogUIEventHandler HideCatalogUIEvent;

		#endregion
	}
}
