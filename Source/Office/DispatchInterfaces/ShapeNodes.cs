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
    /// DispatchInterface ShapeNodes 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
    [Duplicate("NetOffice.ExcelApi.ShapeNodes")]
	[TypeId("000C0319-0000-0000-C000-000000000046")]
    public interface ShapeNodes : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.ShapeNode>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

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
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.ShapeNode this[object index] { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Delete(Int32 index);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
        /// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        /// <param name="x2">optional Single X2 = 0</param>
        /// <param name="y2">optional Single Y2 = 0</param>
        /// <param name="x3">optional Single X3 = 0</param>
        /// <param name="y3">optional Single Y3 = 0</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, object x2, object y2, object x3, object y3);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
        /// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
        /// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        /// <param name="x2">optional Single X2 = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, object x2);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
        /// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        /// <param name="x2">optional Single X2 = 0</param>
        /// <param name="y2">optional Single Y2 = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, object x2, object y2);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
        /// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        /// <param name="x2">optional Single X2 = 0</param>
        /// <param name="y2">optional Single Y2 = 0</param>
        /// <param name="x3">optional Single X3 = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, object x2, object y2, object x3);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void SetEditingType(Int32 index, NetOffice.OfficeApi.Enums.MsoEditingType editingType);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void SetPosition(Int32 index, Single x1, Single y1);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void SetSegmentType(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.ShapeNode>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.ShapeNode> GetEnumerator();

        #endregion
    }
}
