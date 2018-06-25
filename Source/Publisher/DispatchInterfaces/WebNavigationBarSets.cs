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
	/// DispatchInterface WebNavigationBarSets 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("A05DB779-1FC2-4288-A150-939A955696C4")]
	public interface WebNavigationBarSets : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.WebNavigationBarSet>
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
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.WebNavigationBarSet this[object index] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="design">optional NetOffice.PublisherApi.Enums.PbWizardNavBarDesign Design = 1</param>
		/// <param name="autoUpdate">optional bool AutoUpdate = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.WebNavigationBarSet AddSet(string name, object design, object autoUpdate);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.WebNavigationBarSet AddSet(string name);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="design">optional NetOffice.PublisherApi.Enums.PbWizardNavBarDesign Design = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.WebNavigationBarSet AddSet(string name, object design);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.WebNavigationBarSet>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.WebNavigationBarSet> GetEnumerator();

        #endregion
    }
}
