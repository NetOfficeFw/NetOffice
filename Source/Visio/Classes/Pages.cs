using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Pages_PageAddedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Pages_PageChangedEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Pages_BeforePageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Pages_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Pages_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Pages_ConnectionsAddedEventHandler(NetOffice.VisioApi.IVConnects cnnects);
	public delegate void Pages_ConnectionsDeletedEventHandler(NetOffice.VisioApi.IVConnects connects);
	public delegate void Pages_QueryCancelPageDeleteEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Pages_PageDeleteCanceledEventHandler(NetOffice.VisioApi.IVPage page);
	public delegate void Pages_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape sape);
	public delegate void Pages_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Pages_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Pages_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Pages_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Pages_ContainerRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Pages_ContainerRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Pages_CalloutRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Pages_CalloutRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent shapePair);
	public delegate void Pages_QueryCancelReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Pages_ReplaceShapesCanceledEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Pages_BeforeReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Pages_AfterReplaceShapesEventHandler(NetOffice.VisioApi.IVSelection sel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Pages 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769367(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EPages))]
	[TypeId("000D0A05-0000-0000-C000-000000000046")]
    public interface Pages : IVPages, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766451(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_PageAddedEventHandler PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767106(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_PageChangedEventHandler PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766361(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_BeforePageDeleteEventHandler BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767033(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765740(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767610(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766978(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768792(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766559(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768852(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767501(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767051(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ConnectionsAddedEventHandler ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767866(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ConnectionsDeletedEventHandler ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768301(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766924(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_PageDeleteCanceledEventHandler PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767530(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765811(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768215(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768528(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765905(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767904(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767587(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767504(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767711(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Pages_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765076(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Pages_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765855(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Pages_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768287(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Pages_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766120(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Pages_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768997(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Pages_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767353(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Pages_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768900(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Pages_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766166(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Pages_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766526(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		event Pages_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Pages_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Pages_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Pages_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		event Pages_AfterReplaceShapesEventHandler AfterReplaceShapesEvent;

		#endregion
	}
}
