using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotAxis 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("F5B39B2B-1480-11D3-8549-00C04FAC67D7")]
	public interface PivotAxis : ICOMObject
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
		NetOffice.OWC10Api.PivotFieldSets FieldSets { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotLabel Label { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">NetOffice.OWC10Api.PivotFieldSet fieldSet</param>
		/// <param name="before">optional object before</param>
		/// <param name="remove">optional bool Remove = true</param>
		[SupportByVersion("OWC10", 1)]
		void InsertFieldSet(NetOffice.OWC10Api.PivotFieldSet fieldSet, object before, object remove);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">NetOffice.OWC10Api.PivotFieldSet fieldSet</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void InsertFieldSet(NetOffice.OWC10Api.PivotFieldSet fieldSet);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">NetOffice.OWC10Api.PivotFieldSet fieldSet</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void InsertFieldSet(NetOffice.OWC10Api.PivotFieldSet fieldSet, object before);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">object fieldSet</param>
		[SupportByVersion("OWC10", 1)]
		void RemoveFieldSet(object fieldSet);

		#endregion
	}
}
