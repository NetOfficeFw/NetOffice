using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Documents_DocumentOpenedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_DocumentCreatedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_DocumentSavedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_DocumentSavedAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_DocumentChangedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_BeforeDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_StyleAddedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Documents_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Documents_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Documents_MasterAddedEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Documents_MasterChangedEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Documents_BeforeMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Documents_PageAddedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Documents_PageChangedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Documents_BeforePageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Documents_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Documents_RunModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_DesignModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_BeforeDocumentSaveEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_BeforeDocumentSaveAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Documents_ConnectionsAddedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Documents_ConnectionsDeletedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Documents_QueryCancelDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_DocumentCloseCanceledEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Documents_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Documents_QueryCancelMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Documents_MasterDeleteCanceledEventHandler(NetOffice.VisioApi.IVMaster master);
	public delegate void Documents_QueryCancelPageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Documents_PageDeleteCanceledEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Documents_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Documents_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Documents_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void Documents_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent dataRecordsetChanged);
	public delegate void Documents_DataRecordsetAddedEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void Documents_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Documents_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Documents_AfterRemoveHiddenInformationEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Documents_ContainerRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Documents_ContainerRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Documents_CalloutRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Documents_CalloutRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Documents_RuleSetValidatedEventHandler(NetOffice.VisioApi.IVValidationRuleSet ruleSet);
	public delegate void Documents_QueryCancelReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Documents_ReplaceShapesCanceledEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Documents_BeforeReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Documents_AfterReplaceShapesEventHandler(NetOffice.VisioApi.IVSelection sel);
	public delegate void Documents_AfterDocumentMergeEventHandler(NetOffice.VisioApi.IVCoauthMergeEvent coauthMergeObjects);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Documents 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769272(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EDocuments))]
	[TypeId("000D0A00-0000-0000-C000-000000000046")]
    public interface Documents : IVDocuments, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768065(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DocumentOpenedEventHandler DocumentOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765969(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DocumentCreatedEventHandler DocumentCreatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767193(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DocumentSavedEventHandler DocumentSavedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768614(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DocumentSavedAsEventHandler DocumentSavedAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767380(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DocumentChangedEventHandler DocumentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766591(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeDocumentCloseEventHandler BeforeDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768688(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_StyleAddedEventHandler StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765792(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_StyleChangedEventHandler StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767015(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767785(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_MasterAddedEventHandler MasterAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765685(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_MasterChangedEventHandler MasterChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768747(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeMasterDeleteEventHandler BeforeMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767475(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_PageAddedEventHandler PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765929(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_PageChangedEventHandler PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765416(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforePageDeleteEventHandler BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768529(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765195(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768191(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768506(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767758(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767202(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767568(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767625(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_RunModeEnteredEventHandler RunModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768448(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DesignModeEnteredEventHandler DesignModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767086(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeDocumentSaveEventHandler BeforeDocumentSaveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769043(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeDocumentSaveAsEventHandler BeforeDocumentSaveAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765755(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768568(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ConnectionsAddedEventHandler ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767619(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ConnectionsDeletedEventHandler ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766172(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelDocumentCloseEventHandler QueryCancelDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765090(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_DocumentCloseCanceledEventHandler DocumentCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765860(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767497(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766142(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelMasterDeleteEventHandler QueryCancelMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766242(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_MasterDeleteCanceledEventHandler MasterDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765447(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768576(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_PageDeleteCanceledEventHandler PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768023(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765928(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768072(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769042(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765818(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767895(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767818(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765630(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767667(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Documents_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767452(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768212(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766204(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769192(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765966(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_DataRecordsetChangedEventHandler DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766714(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_DataRecordsetAddedEventHandler DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765797(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765391(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766845(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Documents_AfterRemoveHiddenInformationEventHandler AfterRemoveHiddenInformationEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767337(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Documents_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765397(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Documents_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768274(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Documents_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765507(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Documents_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766503(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Documents_RuleSetValidatedEventHandler RuleSetValidatedEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Documents_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Documents_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Documents_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Documents_AfterReplaceShapesEventHandler AfterReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Documents_AfterDocumentMergeEventHandler AfterDocumentMergeEvent;

		#endregion
	}
}
