using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Paragraph 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845307.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Paragraph : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198036.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Range { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193407.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840156.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845450.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838951.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ParagraphFormat Format { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839415.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TabStops TabStops { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834531.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Borders Borders { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193363.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.DropCap DropCap { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835808.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		object Style { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844837.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdParagraphAlignment Alignment { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835246.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 KeepTogether { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196384.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 KeepWithNext { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821677.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 PageBreakBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840640.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 NoLineNumber { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192362.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single RightIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192031.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LeftIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194707.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single FirstLineIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840537.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LineSpacing { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844903.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLineSpacing LineSpacingRule { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single SpaceBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191990.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single SpaceAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834839.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Hyphenation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196595.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 WidowControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822389.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Shading Shading { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834812.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 FarEastLineBreakControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838937.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 WordWrap { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822614.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 HangingPunctuation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196345.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 HalfWidthPunctuationOnTopOfLine { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194017.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AddSpaceBetweenFarEastAndAlpha { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836922.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AddSpaceBetweenFarEastAndDigit { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192567.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdBaselineAlignment BaseLineAlignment { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192561.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AutoAdjustRightIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821565.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 DisableLineHeightGrid { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197419.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdOutlineLevel OutlineLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840651.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single CharacterUnitRightIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192004.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single CharacterUnitLeftIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844832.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single CharacterUnitFirstLineIndent { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839860.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LineUnitBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845221.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LineUnitAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836291.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdReadingOrder ReadingOrder { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840491.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string ID { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195344.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 SpaceBeforeAuto { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838152.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 SpaceAfterAuto { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198321.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool IsStyleSeparator { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196956.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 MirrorIndents { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822871.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdTextboxTightWrap TextboxTightWrap { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840388.aspx </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_ListNumberOriginal(Int16 level);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Alias for get_ListNumberOriginal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840388.aspx </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16), Redirect("get_ListNumberOriginal")]
		Int16 ListNumberOriginal(Int16 level);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 ParaID { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 TextID { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229749.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		bool CollapsedState { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232325.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		bool CollapseHeadingByDefault { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839929.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CloseUp();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197449.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenUp();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836114.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OpenOrCloseUp();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837340.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void TabHangingIndent(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198348.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void TabIndent(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834877.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Reset();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839871.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Space1();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838089.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Space15();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Space2();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839146.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void IndentCharWidth(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191696.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void IndentFirstLineCharWidth(Int16 count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196530.aspx </remarks>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Paragraph Next(object count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196530.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Paragraph Next();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845423.aspx </remarks>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Paragraph Previous(object count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845423.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Paragraph Previous();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821068.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OutlinePromote();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844908.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OutlineDemote();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194241.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void OutlineDemoteToBody();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196925.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Indent();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192199.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Outdent();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834939.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SelectNumber();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		/// <param name="level8">optional Int16 Level8 = 0</param>
		/// <param name="level9">optional Int16 Level9 = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7, object level8, object level9);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3, object level4);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		/// <param name="level8">optional Int16 Level8 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7, object level8);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196834.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void ResetAdvanceTo();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845535.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void SeparateList();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192823.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void JoinList();

		#endregion
	}
}
