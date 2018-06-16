using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotCell 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39B31-1480-11D3-8549-00C04FAC67D7")]
	public interface PivotCell : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotAggregates Aggregates { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool Expanded { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.ADODBApi.Recordset Recordset { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotRowMember RowMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotColumnMember ColumnMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DetailTop { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotDetailCell get_DetailCells(Int32 row, Int32 column);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailCells
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailCells")]
		NetOffice.OWC10Api.PivotDetailCell DetailCells(Int32 row, Int32 column);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotDetailCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotDetailCell bottomRight</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotDetailRange get_DetailRange(NetOffice.OWC10Api.PivotDetailCell topLeft, NetOffice.OWC10Api.PivotDetailCell bottomRight);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailRange
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotDetailCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotDetailCell bottomRight</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailRange")]
		NetOffice.OWC10Api.PivotDetailRange DetailRange(NetOffice.OWC10Api.PivotDetailCell topLeft, NetOffice.OWC10Api.PivotDetailCell bottomRight);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotData Data { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DetailTopOffset { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DetailRowCount { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DetailColumnCount { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotPageMember PageMember { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailTop">Int32 detailTop</param>
		/// <param name="detailTopOffset">Int32 detailTopOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		void MoveDetailTop(Int32 detailTop, Int32 detailTopOffset, object update);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailTop">Int32 detailTop</param>
		/// <param name="detailTopOffset">Int32 detailTopOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void MoveDetailTop(Int32 detailTop, Int32 detailTopOffset);

		#endregion
	}
}
