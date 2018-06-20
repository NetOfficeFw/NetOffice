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
    /// Interface SeriesCollection 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000C170A-0000-0000-C000-000000000046")]
    public interface SeriesCollection : ICOMObject, IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>
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
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.IMsoSeries this[object index] { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoSeries Add(object source);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels, object categoryLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Extend(object source, object rowcol, object categoryLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Extend(object source);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional object rowcol</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Extend(object source, object rowcol);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="replace">optional object replace</param>
        /// <param name="newSeries">optional object newSeries</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Paste(object rowcol, object seriesLabels, object categoryLabels, object replace, object newSeries);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Paste();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Paste(object rowcol);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Paste(object rowcol, object seriesLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Paste(object rowcol, object seriesLabels, object categoryLabels);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="replace">optional object replace</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object Paste(object rowcol, object seriesLabels, object categoryLabels, object replace);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IMsoSeries NewSeries();

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.IMsoSeries>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.IMsoSeries> GetEnumerator();

        #endregion
    }
}
