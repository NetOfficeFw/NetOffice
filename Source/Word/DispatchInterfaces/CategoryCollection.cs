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
	/// DispatchInterface CategoryCollection 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231100.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Word", 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("04124C2D-039D-4442-9C68-8FA38D11DDD6")]
	public interface CategoryCollection : ICOMObject, IEnumerableProvider<NetOffice.WordApi.ChartCategory>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231978.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231429.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229558.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232266.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Creator { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="index">object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.ChartCategory this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.WordApi.ChartCategory>

        /// <summary>
        /// SupportByVersion Word, 15, 16
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        new IEnumerator<NetOffice.WordApi.ChartCategory> GetEnumerator();

        #endregion
    }
}
