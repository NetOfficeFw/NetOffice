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
	/// Interface IPivotTableChangeList 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000244C1-0001-0000-C000-000000000046")]
	public interface IPivotTableChangeList : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.ValueChange>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.ValueChange this[object index] { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		/// <param name="allocationValue">optional object allocationValue</param>
		/// <param name="allocationMethod">optional object allocationMethod</param>
		/// <param name="allocationWeightExpression">optional object allocationWeightExpression</param>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ValueChange Add(string tuple, Double value, object allocationValue, object allocationMethod, object allocationWeightExpression);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ValueChange Add(string tuple, Double value);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		/// <param name="allocationValue">optional object allocationValue</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ValueChange Add(string tuple, Double value, object allocationValue);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		/// <param name="allocationValue">optional object allocationValue</param>
		/// <param name="allocationMethod">optional object allocationMethod</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ValueChange Add(string tuple, Double value, object allocationValue, object allocationMethod);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.ValueChange>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.ValueChange> GetEnumerator();

        #endregion
    }
}
