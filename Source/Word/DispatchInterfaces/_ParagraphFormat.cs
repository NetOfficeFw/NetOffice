using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface _ParagraphFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	public interface _ParagraphFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821435.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195619.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840947.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838308.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ParagraphFormat Duplicate { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821948.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		object Style { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197202.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdParagraphAlignment Alignment { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821554.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 KeepTogether { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196927.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 KeepWithNext { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836582.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 PageBreakBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195075.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 NoLineNumber { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839317.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single RightIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837464.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LeftIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836045.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single FirstLineIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193108.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LineSpacing { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835443.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLineSpacing LineSpacingRule { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839098.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single SpaceBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821851.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single SpaceAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191759.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Hyphenation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194780.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 WidowControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196025.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 FarEastLineBreakControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839113.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 WordWrap { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835163.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 HangingPunctuation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840529.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 HalfWidthPunctuationOnTopOfLine { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193427.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AddSpaceBetweenFarEastAndAlpha { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834823.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AddSpaceBetweenFarEastAndDigit { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822309.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdBaselineAlignment BaseLineAlignment { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821247.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AutoAdjustRightIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822938.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 DisableLineHeightGrid { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835206.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TabStops TabStops { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838962.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Borders Borders { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192176.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Shading Shading { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839401.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdOutlineLevel OutlineLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840234.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single CharacterUnitRightIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836968.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single CharacterUnitLeftIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840585.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single CharacterUnitFirstLineIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845064.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LineUnitBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192386.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LineUnitAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195163.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdReadingOrder ReadingOrder { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837886.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 SpaceBeforeAuto { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837928.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 SpaceAfterAuto { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195675.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 MirrorIndents { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845204.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdTextboxTightWrap TextboxTightWrap { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229454.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 CollapsedByDefault { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844885.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CloseUp();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845804.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenUp();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821567.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenOrCloseUp();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834264.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void TabHangingIndent(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839143.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void TabIndent(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837308.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Reset();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196259.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Space1();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197455.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Space15();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198344.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Space2();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195772.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void IndentCharWidth(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834550.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void IndentFirstLineCharWidth(Int16 count);

		#endregion
	}
}
