using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface ICalculatedMembers 
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("00024454-0001-0000-C000-000000000046")]
    public interface ICalculatedMembers : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.ExcelApi.CalculatedMember this[object index] { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember Add(string name, string formula, object solveOrder, object type);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="dynamic">optional object dynamic</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="hierarchizeDistinct">optional object hierarchizeDistinct</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic, object displayFolder, object hierarchizeDistinct);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember Add(string name, string formula);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember Add(string name, string formula, object solveOrder);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="dynamic">optional object dynamic</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="dynamic">optional object dynamic</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic, object displayFolder);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula, object solveOrder, object type);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula, object solveOrder);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        /// <param name="parentHierarchy">optional object parentHierarchy</param>
        /// <param name="parentMember">optional object parentMember</param>
        /// <param name="numberFormat">optional object numberFormat</param>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy, object parentMember, object numberFormat);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        /// <param name="parentHierarchy">optional object parentHierarchy</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        /// <param name="parentHierarchy">optional object parentHierarchy</param>
        /// <param name="parentMember">optional object parentMember</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy, object parentMember);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.CalculatedMember>

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.CalculatedMember> GetEnumerator();

        #endregion
    }
}
