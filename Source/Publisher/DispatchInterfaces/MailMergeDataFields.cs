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
	/// DispatchInterface MailMergeDataFields 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("9DF97ACB-FE7B-47E9-A644-62F8F663C5D4")]
	public interface MailMergeDataFields : NetOffice.OfficeApi._IMsoDispObj, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.MailMergeDataField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="varIndex">object varIndex</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.MailMergeDataField this[object varIndex] { get; }

        #endregion


        #region IEnumerable<NetOffice.PublisherApi.MailMergeDataField>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.PublisherApi.MailMergeDataField> GetEnumerator();

        #endregion
    }
}
