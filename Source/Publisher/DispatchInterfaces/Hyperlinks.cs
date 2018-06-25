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
	/// DispatchInterface Hyperlinks 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("72E5D393-FC3B-4A38-A9DA-AA30370E66F9")]
	public interface Hyperlinks : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Hyperlink>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Hyperlink this[Int32 index] { get; }

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
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 1</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		/// <param name="textToDisplay">optional string TextToDisplay = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address, object relativePage, object pageID, object textToDisplay);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address, object relativePage);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 1</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address, object relativePage, object pageID);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Hyperlink>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Hyperlink> GetEnumerator();

        #endregion
    }
}
