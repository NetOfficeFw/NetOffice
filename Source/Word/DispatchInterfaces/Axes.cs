using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Axes 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196213.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Word", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("354AB591-A217-48B4-99E4-14F58F15667D")]
	public interface Axes : ICOMObject, IEnumerableProvider<NetOffice.WordApi.Axis>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839419.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837689.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845255.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835806.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Custom Indexer
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.XlAxisType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14, 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
		NetOffice.WordApi.Axis this[NetOffice.WordApi.Enums.XlAxisType type] { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.XlAxisType type</param>
		/// <param name="axisGroup">optional NetOffice.WordApi.Enums.XlAxisGroup axisGroup = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.Axis this[NetOffice.WordApi.Enums.XlAxisType type, NetOffice.WordApi.Enums.XlAxisGroup axisGroup] { get; }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Axis>

        /// <summary>
        /// SupportByVersion Word, 14,15,16
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.Axis> GetEnumerator();

        #endregion
    }
}
