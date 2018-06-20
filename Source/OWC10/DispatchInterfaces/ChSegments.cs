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
	/// DispatchInterface ChSegments 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("BEAE3F41-6FB5-11D2-A8FF-00C04FAC6830")]
	public interface ChSegments : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.ChSegment>
	{
		#region Properties

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
		NetOffice.OWC10Api.ChFormatMap Parent { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.ChSegment this[Int32 index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChSegment Add();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("OWC10", 1)]
		void Delete(Int32 index);

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.ChSegment>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OWC10Api.ChSegment> GetEnumerator();

        #endregion
    }
}
