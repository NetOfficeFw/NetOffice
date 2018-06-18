using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Extender_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Extender_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Extender_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Extender_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Extender_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Extender_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Extender 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EShape))]
	[TypeId("000D0D06-0000-0000-C000-000000000046")]
    public interface Extender : IVDispExtender, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Extender_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Extender_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Extender_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Extender_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Extender_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Extender_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		#endregion
	}
}
