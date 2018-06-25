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
	/// DispatchInterface WebListBoxItems 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Value, EnumeratorInvoke.Custom, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("E7924062-8668-11D3-9058-00C04F799E3F")]
	public interface WebListBoxItems : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<string>
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
		/// <param name="item">string item</param>
		/// <param name="index">optional Int32 Index = -1</param>
		/// <param name="selectState">optional bool SelectState = false</param>
		/// <param name="itemValue">optional string itemValue</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void AddItem(string item, object index, object selectState, object itemValue);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="item">string item</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddItem(string item);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="item">string item</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddItem(string item, object index);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="item">string item</param>
		/// <param name="index">optional Int32 Index = -1</param>
		/// <param name="selectState">optional bool SelectState = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddItem(string item, object index, object selectState);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Delete(Int32 index);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		string this[object index] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="selectState">bool selectState</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Selected(Int32 index, bool selectState);

        #endregion

        #region IEnumerable<string> Member

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<string> GetEnumerator();

        #endregion
    }
}
