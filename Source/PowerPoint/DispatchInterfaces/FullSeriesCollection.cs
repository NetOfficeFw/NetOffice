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
	/// DispatchInterface FullSeriesCollection 
	/// SupportByVersion PowerPoint, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227661.aspx </remarks>
	[SupportByVersion("PowerPoint", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "PowerPoint", 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("288B25A9-98EF-41E5-BEBA-F547D7169BF2")]
	public interface FullSeriesCollection : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.Series>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj684164.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227663.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj717740.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj684239.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		Int32 Creator { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Series this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Series>

        /// <summary>
        /// SupportByVersion PowerPoint, 15, 16
        /// </summary>
        [SupportByVersion("PowerPoint", 15, 16)]
        new IEnumerator<NetOffice.PowerPointApi.Series> GetEnumerator();

        #endregion
    }
}
