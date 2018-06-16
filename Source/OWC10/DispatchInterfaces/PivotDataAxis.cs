using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotDataAxis 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39B43-1480-11D3-8549-00C04FAC67D7")]
	public interface PivotDataAxis : PivotAxis
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotTotals Totals { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">NetOffice.OWC10Api.PivotTotal total</param>
		/// <param name="before">optional object before</param>
		[SupportByVersion("OWC10", 1)]
		void InsertTotal(NetOffice.OWC10Api.PivotTotal total, object before);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">NetOffice.OWC10Api.PivotTotal total</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void InsertTotal(NetOffice.OWC10Api.PivotTotal total);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">object total</param>
		[SupportByVersion("OWC10", 1)]
		void RemoveTotal(object total);

		#endregion
	}
}
