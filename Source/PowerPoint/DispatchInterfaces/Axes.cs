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
	/// DispatchInterface Axes 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745246.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "PowerPoint", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("92D41A52-F07E-4CA4-AF6F-BEF486AA4E6F")]
	public interface Axes : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.Axis>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745728.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744181.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744589.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746281.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Custom Indexer
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.XlAxisType type</param>
		[SupportByVersion("PowerPoint", 14, 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
		NetOffice.PowerPointApi.Axis this[NetOffice.PowerPointApi.Enums.XlAxisType type] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.XlAxisType type</param>
		/// <param name="axisGroup">optional NetOffice.PowerPointApi.Enums.XlAxisGroup axisGroup = 1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Axis this[NetOffice.PowerPointApi.Enums.XlAxisType type, NetOffice.PowerPointApi.Enums.XlAxisGroup axisGroup] { get; }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Axis>

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        new IEnumerator<NetOffice.PowerPointApi.Axis> GetEnumerator();

        #endregion
    }
}
