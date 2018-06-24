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
	/// Interface ISlicerCaches 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000244C3-0001-0000-C000-000000000046")]
	public interface ISlicerCaches : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.SlicerCache>
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
		[SupportByVersion("Excel", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.SlicerCache this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="sourceField">object sourceField</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.SlicerCache Add(object source, object sourceField, object name);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="sourceField">object sourceField</param>
		/// <param name="name">optional object name</param>
		/// <param name="slicerCacheType">optional object slicerCacheType</param>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.SlicerCache Add(object source, object sourceField, object name, object slicerCacheType);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="sourceField">object sourceField</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.SlicerCache Add(object source, object sourceField);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="sourceField">object sourceField</param>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.SlicerCache _Add(object source, object sourceField, object name);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="sourceField">object sourceField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.SlicerCache _Add(object source, object sourceField);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.SlicerCache>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.SlicerCache> GetEnumerator();
        
        #endregion
    }
}
