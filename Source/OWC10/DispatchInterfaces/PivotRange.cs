using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotRange 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39B4A-1480-11D3-8549-00C04FAC67D7")]
	public interface PivotRange : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotCell TopLeft { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotCell BottomRight { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotRowMembers RowMembers { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotColumnMembers ColumnMembers { get; }

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

		#endregion

	}
}
