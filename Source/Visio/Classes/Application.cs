using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_AppActivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AppDeactivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AppObjActivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AppObjDeactivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeQuitEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeModalEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AfterModalEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_WindowOpenedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_DocumentOpenedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentCreatedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentSavedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentSavedAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentChangedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_BeforeDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_StyleAddedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Application_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Application_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Application_MasterAddedEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Application_MasterChangedEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Application_BeforeMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Application_PageAddedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Application_PageChangedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Application_BeforePageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Application_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Application_MarkerEventEventHandler(NetOffice.VisioApi.IVApplication app, Int32 sequenceNum, string contextString);
	public delegate void Application_NoEventsPendingEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_VisioIsIdleEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_MustFlushScopeBeginningEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_MustFlushScopeEndedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_RunModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DesignModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_BeforeDocumentSaveEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_BeforeDocumentSaveAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell Cell);
	public delegate void Application_ConnectionsAddedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Application_ConnectionsDeletedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Application_EnterScopeEventHandler(NetOffice.VisioApi.IVApplication app, Int32 nScopeID, string bstrDescription);
	public delegate void Application_ExitScopeEventHandler(NetOffice.VisioApi.IVApplication app, Int32 nScopeID, string bstrDescription, bool bErrOrCancelled);
	public delegate void Application_QueryCancelQuitEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_QuitCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Application_QueryCancelDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentCloseCanceledEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Application_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Application_QueryCancelMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Application_MasterDeleteCanceledEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Application_QueryCancelPageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Application_PageDeleteCanceledEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Application_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_QueryCancelSuspendEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_SuspendCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeSuspendEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AfterResumeEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap msg);
	public delegate void Application_MouseDownEventHandler(Int32 button, Int32 KeyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Application_MouseMoveEventHandler(Int32 button, Int32 KeyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Application_MouseUpEventHandler(Int32 button, Int32 KeyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Application_KeyDownEventHandler(Int32 keyCode, Int32 KeyButtonState, ref bool CancelDefault);
	public delegate void Application_KeyPressEventHandler(Int32 keyAscii, ref bool CancelDefault);
	public delegate void Application_KeyUpEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void Application_QueryCancelSuspendEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_SuspendEventsCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeSuspendEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AfterResumeEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Application_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Application_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void Application_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent dataRecordsetChanged);
	public delegate void Application_DataRecordsetAddedEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void Application_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Application_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Application_AfterRemoveHiddenInformationEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_ContainerRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Application_ContainerRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Application_CalloutRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Application_CalloutRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Application_RuleSetValidatedEventHandler(NetOffice.VisioApi.IVValidationRuleSet ruleSet);
	public delegate void Application_QueryCancelReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Application_ReplaceShapesCanceledEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Application_BeforeReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Application_AfterReplaceShapesEventHandler(NetOffice.VisioApi.IVSelection sel);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.VisioApi.Behind.Application
    /// SupportByVersion Visio 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.VisioApi.Behind.Application
    {
        private string _defaultProgId = "Visio.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Visio
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Visio based on given id.
        /// This can be used to target a specific version of Microsoft Visio.
        /// Example usage:
        /// "Microsoft.Visio.12" to target Visio 2007
        /// "Microsoft.Visio.14" to target Visio 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Visio
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Visio
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
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769220(v=office.14).aspx </remarks>
    [SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Visio.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EApplication))]
	[TypeId("00021A20-0000-0000-C000-000000000046")]
    public interface Application : IVApplication, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765356(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_AppActivatedEventHandler AppActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765903(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_AppDeactivatedEventHandler AppDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767797(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_AppObjActivatedEventHandler AppObjActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765196(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_AppObjDeactivatedEventHandler AppObjDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767832(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeQuitEventHandler BeforeQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766316(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeModalEventHandler BeforeModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768670(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_AfterModalEventHandler AfterModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767725(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_WindowOpenedEventHandler WindowOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768425(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_SelectionChangedEventHandler SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768648(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeWindowClosedEventHandler BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768932(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_WindowActivatedEventHandler WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765921(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768604(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769065(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_WindowTurnedToPageEventHandler WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768552(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DocumentOpenedEventHandler DocumentOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765843(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DocumentCreatedEventHandler DocumentCreatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767627(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DocumentSavedEventHandler DocumentSavedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768947(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DocumentSavedAsEventHandler DocumentSavedAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768119(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DocumentChangedEventHandler DocumentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768151(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeDocumentCloseEventHandler BeforeDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767751(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_StyleAddedEventHandler StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769029(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_StyleChangedEventHandler StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766539(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768930(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MasterAddedEventHandler MasterAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769090(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MasterChangedEventHandler MasterChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766726(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeMasterDeleteEventHandler BeforeMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765378(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_PageAddedEventHandler PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768083(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_PageChangedEventHandler PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766722(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforePageDeleteEventHandler BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766392(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766131(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767789(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766974(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767938(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767908(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767326(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765486(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MarkerEventEventHandler MarkerEventEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767335(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_NoEventsPendingEventHandler NoEventsPendingEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766439(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_VisioIsIdleEventHandler VisioIsIdleEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767512(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MustFlushScopeBeginningEventHandler MustFlushScopeBeginningEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768052(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MustFlushScopeEndedEventHandler MustFlushScopeEndedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765975(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_RunModeEnteredEventHandler RunModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765830(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DesignModeEnteredEventHandler DesignModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768487(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeDocumentSaveEventHandler BeforeDocumentSaveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768752(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeDocumentSaveAsEventHandler BeforeDocumentSaveAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769046(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768100(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ConnectionsAddedEventHandler ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767479(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ConnectionsDeletedEventHandler ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769070(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_EnterScopeEventHandler EnterScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767448(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ExitScopeEventHandler ExitScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765429(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelQuitEventHandler QueryCancelQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765158(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QuitCanceledEventHandler QuitCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765706(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_WindowChangedEventHandler WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765751(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ViewChangedEventHandler ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769020(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765316(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_WindowCloseCanceledEventHandler WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766512(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelDocumentCloseEventHandler QueryCancelDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765332(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_DocumentCloseCanceledEventHandler DocumentCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767117(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768236(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767170(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelMasterDeleteEventHandler QueryCancelMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767357(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MasterDeleteCanceledEventHandler MasterDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767161(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765525(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_PageDeleteCanceledEventHandler PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765842(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768563(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767733(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768575(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766557(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766751(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765728(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765456(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765231(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765467(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_QueryCancelSuspendEventHandler QueryCancelSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766700(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_SuspendCanceledEventHandler SuspendCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766733(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_BeforeSuspendEventHandler BeforeSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766935(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_AfterResumeEventHandler AfterResumeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765211(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769048(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766075(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767334(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766050(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768385(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769131(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Application_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767255(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_QueryCancelSuspendEventsEventHandler QueryCancelSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765864(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_SuspendEventsCanceledEventHandler SuspendEventsCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767714(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_BeforeSuspendEventsEventHandler BeforeSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768214(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_AfterResumeEventsEventHandler AfterResumeEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767917(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768118(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765725(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767894(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767322(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_DataRecordsetChangedEventHandler DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765105(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_DataRecordsetAddedEventHandler DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765626(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768165(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767810(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Application_AfterRemoveHiddenInformationEventHandler AfterRemoveHiddenInformationEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767352(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Application_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765445(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Application_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769019(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Application_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766994(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Application_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768391(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Application_RuleSetValidatedEventHandler RuleSetValidatedEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Application_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Application_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Application_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Application_AfterReplaceShapesEventHandler AfterReplaceShapesEvent;

		#endregion
	}
}
