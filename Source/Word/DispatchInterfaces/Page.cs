using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Page 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193865.aspx </remarks>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("352840A9-AF7D-4CA4-87FC-21C68FDAB3E4")]
	public interface Page : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821074.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835228.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193350.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197564.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Left { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844862.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Top { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195782.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Width { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840957.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 Height { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196254.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Rectangles Rectangles { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845760.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.Breaks Breaks { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821585.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		object EnhMetaFileBits { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 15, 16)]
		void SaveAsPNG(string fileName);

		#endregion
	}
}
