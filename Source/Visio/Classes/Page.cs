using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Page_PageChangedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Page_BeforePageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Page_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_CellChangedEventHandler(NetOffice.VisioApi.IVCell cCell);
	public delegate void Page_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Page_ConnectionsAddedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Page_ConnectionsDeletedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Page_QueryCancelPageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Page_PageDeleteCanceledEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Page_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Page_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Page_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Page_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Page_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dtaRowID);
	public delegate void Page_ContainerRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Page_ContainerRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Page_CalloutRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Page_CalloutRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Page_QueryCancelReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Page_ReplaceShapesCanceledEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Page_BeforeReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Page_AfterReplaceShapesEventHandler(NetOffice.VisioApi.IVSelection sel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Page 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769363(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EPage))]
	[TypeId("000D0A06-0000-0000-C000-000000000046")]
    public interface Page : IVPage, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768718(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_PageChangedEventHandler PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766299(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_BeforePageDeleteEventHandler BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768077(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765432(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768338(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765629(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766984(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768194(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767008(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765978(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766580(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ConnectionsAddedEventHandler ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767066(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ConnectionsDeletedEventHandler ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769084(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766538(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_PageDeleteCanceledEventHandler PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766718(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767341(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769156(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767216(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766237(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767794(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765963(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767759(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768207(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Page_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768916(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Page_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767856(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Page_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768051(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Page_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766034(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Page_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768669(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Page_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766275(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Page_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765745(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Page_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767961(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Page_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765135(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Page_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Page_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Page_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Page_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Page_AfterReplaceShapesEventHandler AfterReplaceShapesEvent;

		#endregion
	}
}
