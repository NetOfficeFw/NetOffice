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
	/// DispatchInterface FieldListDragDataList 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("2A9DDE7C-D83E-11D3-AE6C-00C04F613171")]
	public interface FieldListDragDataList : ICOMObject, IEnumerableProvider<NetOffice.OWC10Api.FieldListDragData>
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
		/// <param name="index">Int32 index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OWC10Api.FieldListDragData this[Int32 index] { get; }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.FieldListDragData>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OWC10Api.FieldListDragData> GetEnumerator();

        #endregion
    }
}
