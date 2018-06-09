using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface TextFrame2 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822136.aspx </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.TextFrame2")]
	[TypeId("000C0398-0000-0000-C000-000000000046")]
	public interface TextFrame2 : NetOffice.OfficeApi._IMsoDispObj
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837835.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196549.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single MarginBottom { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194372.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single MarginLeft { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839217.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single MarginRight { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196506.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single MarginTop { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823049.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194998.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoHorizontalAnchor HorizontalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821633.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoVerticalAnchor VerticalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194411.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPathFormat PathFormat { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195004.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoWarpFormat WarpFormat { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822124.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetTextEffect WordArtformat { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840619.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState WordWrap { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835575.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoAutoSize AutoSize { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837830.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.ThreeDFormat ThreeD { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838370.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState HasText { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196881.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.TextRange2 TextRange { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198210.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.TextColumn2 Column { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838834.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Ruler2 Ruler { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840575.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState NoTextRotation { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840426.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		void DeleteText();

		#endregion
	}
}
