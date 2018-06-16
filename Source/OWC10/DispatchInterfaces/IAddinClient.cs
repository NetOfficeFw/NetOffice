using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface IAddinClient 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("198924BD-4102-4CB0-B7E8-DBF8BE7EB5A1")]
	public interface IAddinClient : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="vardisp">object vardisp</param>
		[SupportByVersion("OWC10", 1)]
		void GrantAddinHost(object vardisp);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void RemoveAddinHost();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		/// <param name="semiCalced">bool semiCalced</param>
		[SupportByVersion("OWC10", 1)]
		void IsSemiCalced(Int32 dispid, bool semiCalced);

		#endregion
	}
}
