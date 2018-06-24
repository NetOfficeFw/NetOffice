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
	/// DispatchInterface SlicerCaches 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839595.aspx </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 14, 15, 1), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000244C3-0000-0000-C000-000000000046")]
	public interface SlicerCaches : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.SlicerCache>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197154.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840374.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839797.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194167.aspx </remarks>
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
