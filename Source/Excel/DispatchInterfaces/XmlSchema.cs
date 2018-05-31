using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface XmlSchema 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821806.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface XmlSchema : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839037.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839428.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834702.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840490.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.XmlNamespace Namespace { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194106.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string XML { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835601.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string Name { get; }

		#endregion

	}
}
