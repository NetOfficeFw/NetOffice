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
    /// DispatchInterface PropertyTests 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0334-0000-0000-C000-000000000046")]
    public interface PropertyTests : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PropertyTest>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.PropertyTest this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        /// <param name="value">optional object value</param>
        /// <param name="secondValue">optional object secondValue</param>
        /// <param name="connector">optional NetOffice.OfficeApi.Enums.MsoConnector Connector = 1</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition, object value, object secondValue, object connector);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        /// <param name="value">optional object value</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition, object value);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="condition">NetOffice.OfficeApi.Enums.MsoCondition condition</param>
        /// <param name="value">optional object value</param>
        /// <param name="secondValue">optional object secondValue</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Add(string name, NetOffice.OfficeApi.Enums.MsoCondition condition, object value, object secondValue);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Remove(Int32 index);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PropertyTest>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.PropertyTest> GetEnumerator();

        #endregion
    }
}
