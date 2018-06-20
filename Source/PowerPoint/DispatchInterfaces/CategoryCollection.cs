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
	/// DispatchInterface CategoryCollection 
	/// SupportByVersion PowerPoint, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227558.aspx </remarks>
	[SupportByVersion("PowerPoint", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "PowerPoint", 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("2432F529-514B-4575-AA71-1754C74A13D6")]
	public interface CategoryCollection : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.ChartCategory>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj684258.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228684.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj717694.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj684259.aspx </remarks>
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
		NetOffice.PowerPointApi.ChartCategory this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.ChartCategory>

        /// <summary>
        /// SupportByVersion PowerPoint, 15, 16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("PowerPoint", 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.PowerPointApi.ChartCategory> GetEnumerator();

        #endregion
    }
}
