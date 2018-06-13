using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface XMLChildNodeSuggestion 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("A87E00E9-3AC3-4B53-ABE3-7379653D0E82")]
	public interface XMLChildNodeSuggestion : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		string BaseName { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		string NamespaceURI { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLSchemaReference XMLSchemaReference { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNode Insert(object range);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNode Insert();

		#endregion
	}
}
