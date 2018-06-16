using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotData 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39B30-1480-11D3-8549-00C04FAC67D7")]
	public interface PivotData : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotView View { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotResultRowAxis RowAxis { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotResultColumnAxis ColumnAxis { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotResultFilterAxis FilterAxis { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotResultDataAxis DataAxis { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotColumnMember Left { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotRowMember Top { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotCell get_Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Cells
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Cells")]
		NetOffice.OWC10Api.PivotCell Cells(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_DetailLeft(NetOffice.OWC10Api.PivotColumnMember column);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
        /// <param name="value">Int32r value</param>
        [SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_DetailLeft(NetOffice.OWC10Api.PivotColumnMember column, Int32 value);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_DetailLeft
		/// </summary>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		[SupportByVersion("OWC10", 1), Redirect("get_DetailLeft")]
		Int32 DetailLeft(NetOffice.OWC10Api.PivotColumnMember column);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotCell bottomRight</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotRange get_Range(NetOffice.OWC10Api.PivotCell topLeft, NetOffice.OWC10Api.PivotCell bottomRight);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="topLeft">NetOffice.OWC10Api.PivotCell topLeft</param>
		/// <param name="bottomRight">NetOffice.OWC10Api.PivotCell bottomRight</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		NetOffice.OWC10Api.PivotRange Range(NetOffice.OWC10Api.PivotCell topLeft, NetOffice.OWC10Api.PivotCell bottomRight);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Left2 { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Top2 { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotResultLabel Label { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.IPivotControl Control { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotRowMembers RowMembers { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotColumnMembers ColumnMembers { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotCell CurrentCell { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 LeftOffset { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 TopOffset { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 ViewportTop { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 ViewportLeft { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		/// <param name="page">NetOffice.OWC10Api.PivotPageMember page</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotCell get_CellsEx(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column, NetOffice.OWC10Api.PivotPageMember page);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_CellsEx
		/// </summary>
		/// <param name="row">NetOffice.OWC10Api.PivotRowMember row</param>
		/// <param name="column">NetOffice.OWC10Api.PivotColumnMember column</param>
		/// <param name="page">NetOffice.OWC10Api.PivotPageMember page</param>
		[SupportByVersion("OWC10", 1), Redirect("get_CellsEx")]
		NetOffice.OWC10Api.PivotCell CellsEx(NetOffice.OWC10Api.PivotRowMember row, NetOffice.OWC10Api.PivotColumnMember column, NetOffice.OWC10Api.PivotPageMember page);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotResultPageAxis PageAxis { get; }

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
		bool IsConsistent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="top">NetOffice.OWC10Api.PivotRowMember top</param>
		/// <param name="topOffset">Int32 topOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		void MoveTop(NetOffice.OWC10Api.PivotRowMember top, Int32 topOffset, object update);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="top">NetOffice.OWC10Api.PivotRowMember top</param>
		/// <param name="topOffset">Int32 topOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void MoveTop(NetOffice.OWC10Api.PivotRowMember top, Int32 topOffset);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">NetOffice.OWC10Api.PivotColumnMember left</param>
		/// <param name="leftOffset">Int32 leftOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		void MoveLeft(NetOffice.OWC10Api.PivotColumnMember left, Int32 leftOffset, object update);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">NetOffice.OWC10Api.PivotColumnMember left</param>
		/// <param name="leftOffset">Int32 leftOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void MoveLeft(NetOffice.OWC10Api.PivotColumnMember left, Int32 leftOffset);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void ShowDetails();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void HideDetails();

		#endregion
	}
}
