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
	/// DispatchInterface Attachments 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("919317EA-D59A-492E-9860-8B45E82CEC0C")]
	public interface Attachments : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Attachment>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="item">Int32 item</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Attachment this[Int32 item] { get; }

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
		/// <param name="filename">string filename</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Attachment Add(string filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void ClearAll();

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Attachment>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Attachment> GetEnumerator();

        #endregion
    }
}
