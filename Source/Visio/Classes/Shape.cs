using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Shape_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Shape_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_TextChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Shape_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection selection);
	public delegate void Shape_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape shape);
	public delegate void Shape_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	public delegate void Shape_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape shape, Int32 dataRecordsetID, Int32 dataRowID);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Shape 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769391(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EShape))]
	[TypeId("000D0A0A-0000-0000-C000-000000000046")]
    public interface Shape : IVShape, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768439(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767288(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_ShapeAddedEventHandler ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765958(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766015(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_ShapeChangedEventHandler ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768302(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_SelectionAddedEventHandler SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766835(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768749(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_TextChangedEventHandler TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768370(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_FormulaChangedEventHandler FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767923(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_ShapeParentChangedEventHandler ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769047(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768049(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768386(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765289(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768619(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_QueryCancelUngroupEventHandler QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767820(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_UngroupCanceledEventHandler UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765418(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769035(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Shape_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767645(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Shape_QueryCancelGroupEventHandler QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767284(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Shape_GroupCanceledEventHandler GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766823(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Shape_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766506(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Shape_ShapeLinkAddedEventHandler ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767434(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event Shape_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent;

		#endregion
	}
}
