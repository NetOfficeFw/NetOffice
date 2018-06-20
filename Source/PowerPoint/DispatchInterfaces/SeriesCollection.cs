using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface SeriesCollection 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745059.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "PowerPoint", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("92D41A76-F07E-4CA4-AF6F-BEF486AA4E6F")]
	public interface SeriesCollection : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.Series>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746728.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744885.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743994.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746247.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746725.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		object Extend(object source, object rowcol, object categoryLabels);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746725.aspx </remarks>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		object Extend(object source);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746725.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		object Extend(object source, object rowcol);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744556.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Series NewSeries();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744310.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744310.aspx </remarks>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Series Add(object source);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744310.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Series Add(object source, object rowcol);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744310.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744310.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Series this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Series>

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        new IEnumerator<NetOffice.PowerPointApi.Series> GetEnumerator();

        #endregion
    }
}
