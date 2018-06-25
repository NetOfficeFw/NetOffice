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
	/// DispatchInterface Documents 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("12F564D6-BC5F-410E-ACD1-7D3D47F06FFA")]
	public interface Documents : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Document>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="varDocument">object varDocument</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Document this[object varDocument] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pbWizard">optional NetOffice.PublisherApi.Enums.PbWizard PbWizard = 0</param>
		/// <param name="desid">optional Int32 desid = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Document Add(object pbWizard, object desid);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Document Add();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pbWizard">optional NetOffice.PublisherApi.Enums.PbWizard PbWizard = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Document Add(object pbWizard);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Document>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Document> GetEnumerator();

        #endregion
    }
}
