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
	/// Interface ISlicers 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000244C7-0001-0000-C000-000000000046")]
	public interface ISlicers : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.Slicer>
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
		NetOffice.ExcelApi.Slicer this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width, object height);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Slicer>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.Slicer> GetEnumerator();

        #endregion
    }
}
