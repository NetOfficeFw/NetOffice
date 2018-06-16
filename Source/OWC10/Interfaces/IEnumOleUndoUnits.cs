using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IEnumOleUndoUnits 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("B3E7C340-EF97-11CE-9BC9-00AA00608E01")]
	public interface IEnumOleUndoUnits : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		/// <param name="rgElt">NetOffice.OWC10Api.IOleUndoUnit rgElt</param>
		/// <param name="pcEltFetched">Int32 pcEltFetched</param>
		[SupportByVersion("OWC10", 1)]
		Int32 RemoteNext(Int32 cElt, out NetOffice.OWC10Api.IOleUndoUnit rgElt, out Int32 pcEltFetched);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Skip(Int32 cElt);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Reset();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Clone(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum);

		#endregion
	}
}
