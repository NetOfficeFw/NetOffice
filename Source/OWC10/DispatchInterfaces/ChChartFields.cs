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
	/// DispatchInterface ChChartFields 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("BB4C16FA-6BEC-11D3-A18A-00C04F612970")]
	public interface ChChartFields : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.ChChartField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.ChChartField this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChDropZone Parent { get; }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.ChChartField> Member

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OWC10Api.ChChartField> GetEnumerator();

        #endregion
    }
}
