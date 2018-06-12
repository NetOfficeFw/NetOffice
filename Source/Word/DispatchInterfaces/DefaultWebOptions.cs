using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface DefaultWebOptions 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820915.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface DefaultWebOptions : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193113.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835798.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195492.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838112.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool OptimizeForBrowser { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845875.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdBrowserLevel BrowserLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839636.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool RelyOnCSS { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193127.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool OrganizeInFolder { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840770.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool UpdateLinksOnSave { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821245.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool UseLongFileNames { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195960.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool CheckIfOfficeIsHTMLEditor { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835154.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool CheckIfWordIsDefaultHTMLEditor { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836595.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool RelyOnVML { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193694.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool AllowPNG { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192205.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoScreenSize ScreenSize { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837317.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 PixelsPerInch { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192595.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoEncoding Encoding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839112.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool AlwaysSaveInDefaultEncoding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193853.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.WebPageFonts Fonts { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193452.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string FolderSuffix { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839638.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTargetBrowser TargetBrowser { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835514.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool SaveNewWebPagesAsWebArchives { get; set; }

		#endregion

	}
}
