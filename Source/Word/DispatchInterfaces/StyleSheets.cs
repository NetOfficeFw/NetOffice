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
	/// DispatchInterface StyleSheets 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195288.aspx </remarks>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("07B7CC7E-E66C-11D3-9454-00105AA31A08")]
	public interface StyleSheets : ICOMObject, IEnumerableProvider<NetOffice.WordApi.StyleSheet>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835790.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836262.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844884.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838761.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.StyleSheet this[object index] { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821965.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkType">NetOffice.WordApi.Enums.WdStyleSheetLinkType linkType</param>
		/// <param name="title">string title</param>
		/// <param name="precedence">NetOffice.WordApi.Enums.WdStyleSheetPrecedence precedence</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.StyleSheet Add(string fileName, NetOffice.WordApi.Enums.WdStyleSheetLinkType linkType, string title, NetOffice.WordApi.Enums.WdStyleSheetPrecedence precedence);

        #endregion

        #region IEnumerable<NetOffice.WordApi.StyleSheet>

        /// <summary>
        /// SupportByVersion Word, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.StyleSheet> GetEnumerator();

        #endregion
    }
}
