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
    /// DispatchInterface SearchScopes 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865267.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0366-0000-0000-C000-000000000046")]
    public interface SearchScopes : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.SearchScope>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.SearchScope this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862063.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SearchScope>

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.SearchScope> GetEnumerator();

        #endregion
    }
}
