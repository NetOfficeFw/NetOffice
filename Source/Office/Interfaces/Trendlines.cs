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
    /// Interface Trendlines 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000C1722-0000-0000-C000-000000000046")]
    public interface Trendlines : ICOMObject, IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

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
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.IMsoTrendline this[object index] { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        /// <param name="displayEquation">optional object displayEquation</param>
        /// <param name="displayRSquared">optional object displayRSquared</param>
        /// <param name="name">optional object name</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        /// <param name="displayEquation">optional object displayEquation</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        /// <param name="displayEquation">optional object displayEquation</param>
        /// <param name="displayRSquared">optional object displayRSquared</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared);

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.IMsoTrendline> GetEnumerator();

        #endregion
    }
}
