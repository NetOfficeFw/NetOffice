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
	/// DispatchInterface PivotAxisMembers 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("30971540-EEB8-11D2-8F35-00600893B533")]
	public interface PivotAxisMembers : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.PivotAxisMember>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.PivotAxisMember this[object index] { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Count { get; }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.PivotAxisMember>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.PivotAxisMember> GetEnumerator();

        #endregion
    }
}
