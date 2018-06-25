using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface TabStops 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("37B3B0B0-44B5-11D3-B65B-00C04F8EF32D")]
	public interface TabStops : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.TabStop>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.TabStop this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void ClearAll();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="position">object position</param>
		/// <param name="alignment">NetOffice.PublisherApi.Enums.PbTabAlignmentType alignment</param>
		/// <param name="leader">NetOffice.PublisherApi.Enums.PbTabLeaderType leader</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Add(object position, NetOffice.PublisherApi.Enums.PbTabAlignmentType alignment, NetOffice.PublisherApi.Enums.PbTabLeaderType leader);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.TabStop>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.PublisherApi.TabStop> GetEnumerator();

        #endregion
    }
}
