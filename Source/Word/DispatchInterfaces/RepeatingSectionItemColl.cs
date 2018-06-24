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
	/// DispatchInterface RepeatingSectionItemColl 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227247.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("53FACA33-DB22-473F-BB51-96C2C86C9304")]
	public interface RepeatingSectionItemColl : ICOMObject, IEnumerableProvider<NetOffice.WordApi.RepeatingSectionItem>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228638.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230383.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229461.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229731.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.RepeatingSectionItem this[Int32 index] { get; }

        #endregion

        #region IEnumerable<NetOffice.WordApi.RepeatingSectionItem> Member

        /// <summary>
        /// SupportByVersion Word, 15, 16
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        new IEnumerator<NetOffice.WordApi.RepeatingSectionItem> GetEnumerator();

        #endregion
    }
}
