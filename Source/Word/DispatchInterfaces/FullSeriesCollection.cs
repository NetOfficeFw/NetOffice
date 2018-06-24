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
	/// DispatchInterface FullSeriesCollection 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231236.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Word", 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("4DACC469-630B-457E-9C8F-08158D57FC7C")]
	public interface FullSeriesCollection : ICOMObject, IEnumerableProvider<NetOffice.WordApi.Series>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230202.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227348.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231922.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227448.aspx </remarks>
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
		NetOffice.WordApi.Series this[object index] { get; }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Series>

        /// <summary>
        /// SupportByVersion Word, 15, 16
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        new IEnumerator<NetOffice.WordApi.Series> GetEnumerator();

        #endregion
    }
}
