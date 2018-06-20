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
	/// DispatchInterface Trendlines 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745537.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "PowerPoint", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("92D41A7A-F07E-4CA4-AF6F-BEF486AA4E6F")]
	public interface Trendlines : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.Trendline>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746317.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746544.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745141.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744597.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746402.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Trendline this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Trendline>

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        new IEnumerator<NetOffice.PowerPointApi.Trendline> GetEnumerator();

        #endregion
    }
}
