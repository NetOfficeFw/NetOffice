using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotClassFactory 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("73F4D511-C851-11D2-8F2D-00600893B533")]
	public interface PivotClassFactory : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell detailCell</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewDetailCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="detailCell">NetOffice.OWC10Api.PivotDetailCell detailCell</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewDetailCell")]
		object NewDetailCell(NetOffice.OWC10Api.PivotDetailCell detailCell);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate aggregate</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewAggregate
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="aggregate">NetOffice.OWC10Api.PivotAggregate aggregate</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewAggregate")]
		object NewAggregate(NetOffice.OWC10Api.PivotAggregate aggregate);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember rowMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewRowMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="rowMember">NetOffice.OWC10Api.PivotRowMember rowMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewRowMember")]
		object NewRowMember(NetOffice.OWC10Api.PivotRowMember rowMember);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember columnMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewColumnMember
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="columnMember">NetOffice.OWC10Api.PivotColumnMember columnMember</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewColumnMember")]
		object NewColumnMember(NetOffice.OWC10Api.PivotColumnMember columnMember);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell cell</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_NewCell(NetOffice.OWC10Api.PivotCell cell);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_NewCell
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="cell">NetOffice.OWC10Api.PivotCell cell</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_NewCell")]
		object NewCell(NetOffice.OWC10Api.PivotCell cell);

		#endregion

	}
}
