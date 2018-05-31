using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface SpellingOptions 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196915.aspx </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface SpellingOptions : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196429.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 DictLang { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835269.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		string UserDict { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194482.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool IgnoreCaps { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840970.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool SuggestMainOnly { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822188.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool IgnoreMixedDigits { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196363.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool IgnoreFileNames { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820788.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool GermanPostReform { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836839.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool KoreanCombineAux { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837047.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool KoreanUseAutoChangeList { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838858.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool KoreanProcessCompound { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838254.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlHebrewModes HebrewModes { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193603.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlArabicModes ArabicModes { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193794.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool ArabicStrictAlefHamza { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835890.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool ArabicStrictFinalYaa { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841131.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool ArabicStrictTaaMarboota { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834631.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool RussianStrictE { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193307.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlSpanishModes SpanishModes { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822381.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlPortugueseReform PortugalReform { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839055.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlPortugueseReform BrazilReform { get; set; }

		#endregion

	}
}
