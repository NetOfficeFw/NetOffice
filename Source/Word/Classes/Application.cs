using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_StartupEventHandler();
	public delegate void Application_QuitEventHandler();
	public delegate void Application_DocumentChangeEventHandler();
	public delegate void Application_DocumentOpenEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_DocumentBeforeCloseEventHandler(NetOffice.WordApi.Document doc, ref bool cancel);
	public delegate void Application_DocumentBeforePrintEventHandler(NetOffice.WordApi.Document doc, ref bool cancel);
	public delegate void Application_DocumentBeforeSaveEventHandler(NetOffice.WordApi.Document doc, ref bool saveAsUI, ref bool cancel);
	public delegate void Application_NewDocumentEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_WindowActivateEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);
	public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.WordApi.Selection sel);
	public delegate void Application_WindowBeforeRightClickEventHandler(NetOffice.WordApi.Selection sel, ref bool cancel);
	public delegate void Application_WindowBeforeDoubleClickEventHandler(NetOffice.WordApi.Selection sel, ref bool cancel);
	public delegate void Application_EPostagePropertyDialogEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_EPostageInsertEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeAfterMergeEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult);
	public delegate void Application_MailMergeAfterRecordMergeEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeBeforeMergeEventHandler(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, ref bool cancel);
	public delegate void Application_MailMergeBeforeRecordMergeEventHandler(NetOffice.WordApi.Document doc, ref bool cancel);
	public delegate void Application_MailMergeDataSourceLoadEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeDataSourceValidateEventHandler(NetOffice.WordApi.Document doc, ref bool handled);
	public delegate void Application_MailMergeWizardSendToCustomEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeWizardStateChangeEventHandler(NetOffice.WordApi.Document doc, ref Int32 fromState, ref Int32 toState, ref bool handled);
	public delegate void Application_WindowSizeEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);
	public delegate void Application_XMLSelectionChangeEventHandler(NetOffice.WordApi.Selection sel, NetOffice.WordApi.XMLNode oldXMLNode, NetOffice.WordApi.XMLNode newXMLNode, ref Int32 reason);
	public delegate void Application_XMLValidationErrorEventHandler(NetOffice.WordApi.XMLNode xmlNode);
	public delegate void Application_DocumentSyncEventHandler(NetOffice.WordApi.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Application_EPostageInsertExEventHandler(NetOffice.WordApi.Document doc, Int32 cpDeliveryAddrStart, Int32 cpDeliveryAddrEnd, Int32 cpReturnAddrStart, Int32 cpReturnAddrEnd, Int32 xaWidth, Int32 yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, ref bool fCancel);
	public delegate void Application_MailMergeDataSourceValidate2EventHandler(NetOffice.WordApi.Document doc, ref bool Handled);
	public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow, ref bool cancel);
	public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow, Int32 closeReason, ref bool cancel);
	public delegate void Application_ProtectedViewWindowSizeEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.WordApi.Behind.Application
    /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online:  http://msdn.microsoft.com/en-us/en-us/library/office/ff838565.aspx </remarks>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.WordApi.Behind.Application
    {
        private string _defaultProgId = "Word.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Word
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Word based on given id.
        /// This can be used to target a specific version of Microsoft Word.
        /// Example usage:
        /// "Microsoft.Word.12" to target Word 2007
        /// "Microsoft.Word.14" to target Word 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Word
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Word
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
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838565.aspx </remarks>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass), ComProgId("Word.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.WordApi.EventContracts.ApplicationEvents2), typeof(NetOffice.WordApi.EventContracts.ApplicationEvents3), typeof(NetOffice.WordApi.EventContracts.ApplicationEvents4))]
	[TypeId("000209FF-0000-0000-C000-000000000046")]
    public interface Application : _Application, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
    {
        #region Events

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_StartupEventHandler StartupEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194164.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_QuitEventHandler QuitEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822189.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_DocumentChangeEventHandler DocumentChangeEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192207.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_DocumentOpenEventHandler DocumentOpenEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834271.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_DocumentBeforeCloseEventHandler DocumentBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845163.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_DocumentBeforePrintEventHandler DocumentBeforePrintEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838299.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_DocumentBeforeSaveEventHandler DocumentBeforeSaveEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836563.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_NewDocumentEventHandler NewDocumentEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840337.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowActivateEventHandler WindowActivateEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198272.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowDeactivateEventHandler WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192791.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837868.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowBeforeRightClickEventHandler WindowBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840048.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowBeforeDoubleClickEventHandler WindowBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197984.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_EPostagePropertyDialogEventHandler EPostagePropertyDialogEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193389.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_EPostageInsertEventHandler EPostageInsertEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198141.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeAfterMergeEventHandler MailMergeAfterMergeEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198157.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeAfterRecordMergeEventHandler MailMergeAfterRecordMergeEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834588.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeBeforeMergeEventHandler MailMergeBeforeMergeEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838357.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeBeforeRecordMergeEventHandler MailMergeBeforeRecordMergeEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196096.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoadEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193130.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidateEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837009.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeWizardSendToCustomEventHandler MailMergeWizardSendToCustomEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838546.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_MailMergeWizardStateChangeEventHandler MailMergeWizardStateChangeEvent;

        /// <summary>
        /// SupportByVersion Word 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834597.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        event Application_WindowSizeEventHandler WindowSizeEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835495.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Application_XMLSelectionChangeEventHandler XMLSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837452.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Application_XMLValidationErrorEventHandler XMLValidationErrorEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835138.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Application_DocumentSyncEventHandler DocumentSyncEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195087.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Application_EPostageInsertExEventHandler EPostageInsertExEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839145.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Application_MailMergeDataSourceValidate2EventHandler MailMergeDataSourceValidate2Event;

        /// <summary>
        /// SupportByVersion Word 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194483.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent;

        /// <summary>
        /// SupportByVersion Word 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192123.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent;

        /// <summary>
        /// SupportByVersion Word 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194718.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Word 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836722.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        event Application_ProtectedViewWindowSizeEventHandler ProtectedViewWindowSizeEvent;

        /// <summary>
        /// SupportByVersion Word 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836396.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent;

        /// <summary>
        /// SupportByVersion Word 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837500.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent;

        #endregion
    }
}
