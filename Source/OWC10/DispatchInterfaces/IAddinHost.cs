using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface IAddinHost 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("FAA0B9C0-F635-44C7-B825-B805F59B3D66")]
	public interface IAddinHost : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varoper">object varoper</param>
		/// <param name="grbit">NetOffice.OWC10Api.Enums.AddinClientTypeEnum grbit</param>
		[SupportByVersion("OWC10", 1)]
		object CoerceOper(object varoper, NetOffice.OWC10Api.Enums.AddinClientTypeEnum grbit);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object RandOper();

		#endregion
	}
}
