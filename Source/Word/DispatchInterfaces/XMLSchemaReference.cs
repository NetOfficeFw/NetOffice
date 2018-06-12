using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface XMLSchemaReference 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195940.aspx </remarks>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface XMLSchemaReference : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192815.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840385.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822620.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194345.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		string NamespaceURI { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193986.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		string Location { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839405.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836321.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Reload();

		#endregion
	}
}
