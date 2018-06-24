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
	/// Interface IPivotFilters 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("00024484-0001-0000-C000-000000000046")]
	public interface IPivotFilters : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.PivotFilter this[object index] { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		/// <param name="wholeDayFilter">optional object wholeDayFilter</param>
		/// <param name="movingPeriod">optional object movingPeriod</param>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField, object wholeDayFilter, object movingPeriod);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		/// <param name="wholeDayFilter">optional object wholeDayFilter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField, object wholeDayFilter);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.PivotFilter>

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.PivotFilter> GetEnumerator();

        #endregion
    }
}
