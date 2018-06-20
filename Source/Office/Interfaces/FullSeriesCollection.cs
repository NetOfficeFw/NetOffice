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
    /// Interface FullSeriesCollection 
    /// SupportByVersion Office, 15, 16
    /// </summary>
    [SupportByVersion("Office", 15, 16)]
    [EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Office", 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000C1732-0000-0000-C000-000000000046")]
    public interface FullSeriesCollection : ICOMObject, IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 15, 16), ProxyResult]
        object Application { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.IMsoSeries this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.IMsoSeries>

        /// <summary>
        /// SupportByVersion Office, 15, 16
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.IMsoSeries> GetEnumerator();

        #endregion
    }
}
