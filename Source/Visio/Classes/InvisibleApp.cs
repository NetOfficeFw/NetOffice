using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void InvisibleApp_AppActivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_AppDeactivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_AppObjActivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_AppObjDeactivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_BeforeQuitEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_BeforeModalEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_AfterModalEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_WindowOpenedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_DocumentOpenedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_DocumentCreatedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_DocumentSavedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_DocumentSavedAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_DocumentChangedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_BeforeDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_StyleAddedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void InvisibleApp_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void InvisibleApp_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void InvisibleApp_MasterAddedEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void InvisibleApp_MasterChangedEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void InvisibleApp_BeforeMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void InvisibleApp_PageAddedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void InvisibleApp_PageChangedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void InvisibleApp_BeforePageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void InvisibleApp_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void InvisibleApp_MarkerEventEventHandler(NetOffice.VisioApi.IVApplication app, Int32 sequenceNum, string contextString);
	public delegate void InvisibleApp_NoEventsPendingEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_VisioIsIdleEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_MustFlushScopeBeginningEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_MustFlushScopeEndedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_RunModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_DesignModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_BeforeDocumentSaveEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_BeforeDocumentSaveAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void InvisibleApp_ConnectionsAddedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void InvisibleApp_ConnectionsDeletedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void InvisibleApp_EnterScopeEventHandler(NetOffice.VisioApi.IVApplication app, Int32 nScopeID, string bstrDescription);
	public delegate void InvisibleApp_ExitScopeEventHandler(NetOffice.VisioApi.IVApplication app, Int32 nScopeID, string bstrDescription, bool bErrOrCancelled);
	public delegate void InvisibleApp_QueryCancelQuitEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_QuitCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void InvisibleApp_QueryCancelDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_DocumentCloseCanceledEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void InvisibleApp_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void InvisibleApp_QueryCancelMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void InvisibleApp_MasterDeleteCanceledEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void InvisibleApp_QueryCancelPageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void InvisibleApp_PageDeleteCanceledEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void InvisibleApp_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_QueryCancelSuspendEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_SuspendCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_BeforeSuspendEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_AfterResumeEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap msg);
	public delegate void InvisibleApp_MouseDownEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void InvisibleApp_MouseMoveEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void InvisibleApp_MouseUpEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void InvisibleApp_KeyDownEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void InvisibleApp_KeyPressEventHandler(Int32 keyAscii, ref bool cancelDefault);
	public delegate void InvisibleApp_KeyUpEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void InvisibleApp_QueryCancelSuspendEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_SuspendEventsCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_BeforeSuspendEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_AfterResumeEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void InvisibleApp_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void InvisibleApp_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void InvisibleApp_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void InvisibleApp_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent dataRecordsetChanged);
	public delegate void InvisibleApp_DataRecordsetAddedEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void InvisibleApp_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void InvisibleApp_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void InvisibleApp_AfterRemoveHiddenInformationEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void InvisibleApp_ContainerRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void InvisibleApp_ContainerRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void InvisibleApp_CalloutRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void InvisibleApp_CalloutRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void InvisibleApp_RuleSetValidatedEventHandler(NetOffice.VisioApi.IVValidationRuleSet ruleSet);
	public delegate void InvisibleApp_QueryCancelReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void InvisibleApp_ReplaceShapesCanceledEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void InvisibleApp_BeforeReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void InvisibleApp_AfterReplaceShapesEventHandler(NetOffice.VisioApi.IVSelection sel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass InvisibleApp 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769303(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EApplication))]
	[TypeId("000D0A26-0000-0000-C000-000000000046")]
    public interface InvisibleApp : IVInvisibleApp, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767393(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_AppActivatedEventHandler AppActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765522(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_AppDeactivatedEventHandler AppDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768447(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_AppObjActivatedEventHandler AppObjActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765372(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_AppObjDeactivatedEventHandler AppObjDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767922(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeQuitEventHandler BeforeQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767572(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeModalEventHandler BeforeModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766355(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_AfterModalEventHandler AfterModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767410(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_WindowOpenedEventHandler WindowOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766811(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_SelectionChangedEventHandler SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768040(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeWindowClosedEventHandler BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767376(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_WindowActivatedEventHandler WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766033(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767063(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767662(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_WindowTurnedToPageEventHandler WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766381(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DocumentOpenedEventHandler DocumentOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767356(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DocumentCreatedEventHandler DocumentCreatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768368(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DocumentSavedEventHandler DocumentSavedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769114(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DocumentSavedAsEventHandler DocumentSavedAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768511(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DocumentChangedEventHandler DocumentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766731(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeDocumentCloseEventHandler BeforeDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767954(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_StyleAddedEventHandler StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767281(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_StyleChangedEventHandler StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765113(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766331(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MasterAddedEventHandler MasterAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768066(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MasterChangedEventHandler MasterChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767032(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeMasterDeleteEventHandler BeforeMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768709(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_PageAddedEventHandler PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768785(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_PageChangedEventHandler PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768579(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforePageDeleteEventHandler BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767730(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767690(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767622(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769104(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767037(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766913(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766877(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765651(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MarkerEventEventHandler MarkerEventEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766723(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_NoEventsPendingEventHandler NoEventsPendingEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766985(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_VisioIsIdleEventHandler VisioIsIdleEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768314(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MustFlushScopeBeginningEventHandler MustFlushScopeBeginningEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768618(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MustFlushScopeEndedEventHandler MustFlushScopeEndedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766960(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_RunModeEnteredEventHandler RunModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768667(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DesignModeEnteredEventHandler DesignModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768909(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeDocumentSaveEventHandler BeforeDocumentSaveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767685(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeDocumentSaveAsEventHandler BeforeDocumentSaveAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765269(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766595(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ConnectionsAddedEventHandler ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767251(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ConnectionsDeletedEventHandler ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766336(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_EnterScopeEventHandler EnterScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768144(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ExitScopeEventHandler ExitScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768147(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelQuitEventHandler QueryCancelQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766219(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QuitCanceledEventHandler QuitCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768984(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_WindowChangedEventHandler WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766826(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ViewChangedEventHandler ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767013(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766190(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_WindowCloseCanceledEventHandler WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766898(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelDocumentCloseEventHandler QueryCancelDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765949(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_DocumentCloseCanceledEventHandler DocumentCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767280(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768716(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768817(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelMasterDeleteEventHandler QueryCancelMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767706(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MasterDeleteCanceledEventHandler MasterDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768454(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765889(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_PageDeleteCanceledEventHandler PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766348(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766839(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766390(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768063(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769183(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767906(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766810(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765068(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765686(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766235(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_QueryCancelSuspendEventHandler QueryCancelSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766499(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_SuspendCanceledEventHandler SuspendCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769072(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_BeforeSuspendEventHandler BeforeSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765491(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_AfterResumeEventHandler AfterResumeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767010(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768379(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767115(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767397(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767556(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768495(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766229(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event InvisibleApp_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765927(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_QueryCancelSuspendEventsEventHandler QueryCancelSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765481(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_SuspendEventsCanceledEventHandler SuspendEventsCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766566(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_BeforeSuspendEventsEventHandler BeforeSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765853(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_AfterResumeEventsEventHandler AfterResumeEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765459(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765402(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765850(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765236(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768548(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_DataRecordsetChangedEventHandler DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767031(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_DataRecordsetAddedEventHandler DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767994(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765058(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767141(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event InvisibleApp_AfterRemoveHiddenInformationEventHandler AfterRemoveHiddenInformationEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765415(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event InvisibleApp_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766762(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event InvisibleApp_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768554(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event InvisibleApp_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765438(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event InvisibleApp_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766746(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event InvisibleApp_RuleSetValidatedEventHandler RuleSetValidatedEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event InvisibleApp_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event InvisibleApp_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event InvisibleApp_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event InvisibleApp_AfterReplaceShapesEventHandler AfterReplaceShapesEvent;

		#endregion
	}
}
