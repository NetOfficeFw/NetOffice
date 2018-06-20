using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface Axes 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000C1712-0000-0000-C000-000000000046")]
    public interface Axes : ICOMObject, IEnumerableProvider<NetOffice.OfficeApi.IMsoAxis>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
        /// Custom Indexer
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.XlAxisType type</param>
		[SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        NetOffice.OfficeApi.IMsoAxis this[NetOffice.OfficeApi.Enums.XlAxisType type] { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="type">NetOffice.OfficeApi.Enums.XlAxisType type</param>
        /// <param name="axisGroup">optional NetOffice.OfficeApi.Enums.XlAxisGroup axisGroup</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.IMsoAxis this[NetOffice.OfficeApi.Enums.XlAxisType type, object axisGroup] { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.IMsoAxis>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.IMsoAxis> GetEnumerator();

        #endregion
    }
}
