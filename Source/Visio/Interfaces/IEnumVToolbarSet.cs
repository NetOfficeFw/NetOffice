using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface IEnumVToolbarSet 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("000D0263-0000-0000-C000-000000000046")]
	public interface IEnumVToolbarSet : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="celt">Int32 celt</param>
		/// <param name="rgelt">NetOffice.VisioApi.IVToolbarSet rgelt</param>
		/// <param name="pceltFetched">Int32 pceltFetched</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 Next(Int32 celt, out NetOffice.VisioApi.IVToolbarSet rgelt, out Int32 pceltFetched);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="celt">Int32 celt</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 Skip(Int32 celt);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 Reset();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="ppenm">NetOffice.VisioApi.IEnumVToolbarSet ppenm</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 Clone(out NetOffice.VisioApi.IEnumVToolbarSet ppenm);

		#endregion
	}
}
