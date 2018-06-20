using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotResultGroupFields 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("3A3AF091-8C35-4CC7-A1F4-B70D239D773B")]
	public interface PivotResultGroupFields : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.PivotResultGroupField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.PivotResultGroupField this[object index] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.PivotResultGroupField>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.PivotResultGroupField> GetEnumerator();

        #endregion
    }
}
