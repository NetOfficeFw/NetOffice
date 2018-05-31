using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface PictureFormat 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834660.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.PictureFormat")]
	public interface PictureFormat : NetOffice.OfficeApi._IMsoDispObj
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195038.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840787.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single Brightness { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822545.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPictureColorType ColorType { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836497.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single Contrast { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838027.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single CropBottom { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840304.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single CropLeft { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836787.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single CropRight { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837812.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single CropTop { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838789.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 TransparencyColor { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836760.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState TransparentBackground { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195104.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.OfficeApi.Crop Crop { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197185.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void IncrementBrightness(Single increment);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822515.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void IncrementContrast(Single increment);

		#endregion
	}
}
