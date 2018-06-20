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
    /// DispatchInterface ODSOFilters 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865224.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Variant, EnumeratorInvoke.Custom, "Office", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("000C1534-0000-0000-C000-000000000046")]
    public interface ODSOFilters : _IMsoDispObj, IEnumerableProvider<object>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860835.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861525.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        object this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864658.aspx </remarks>
        /// <param name="column">string column</param>
        /// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
        /// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
        /// <param name="bstrCompareTo">optional string bstrCompareTo = </param>
        /// <param name="deferUpdate">optional bool DeferUpdate = false</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction, object bstrCompareTo, object deferUpdate);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864658.aspx </remarks>
        /// <param name="column">string column</param>
        /// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
        /// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864658.aspx </remarks>
        /// <param name="column">string column</param>
        /// <param name="comparison">NetOffice.OfficeApi.Enums.MsoFilterComparison comparison</param>
        /// <param name="conjunction">NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction</param>
        /// <param name="bstrCompareTo">optional string bstrCompareTo = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Add(string column, NetOffice.OfficeApi.Enums.MsoFilterComparison comparison, NetOffice.OfficeApi.Enums.MsoFilterConjunction conjunction, object bstrCompareTo);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860318.aspx </remarks>
        /// <param name="index">Int32 index</param>
        /// <param name="deferUpdate">optional bool DeferUpdate = false</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Delete(Int32 index, object deferUpdate);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860318.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Delete(Int32 index);

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
