using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface View 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822898.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface View : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197224.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821357.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844900.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844848.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdViewType Type { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840409.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool FullScreen { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834875.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Draft { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192196.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowAll { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840713.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowFieldCodes { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192217.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MailMergeDataView { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835476.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Magnifier { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191722.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowFirstLineOnly { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821926.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowFormat { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196289.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Zoom Zoom { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197841.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowObjectAnchors { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836054.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowTextBoundaries { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839975.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowHighlight { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840802.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowDrawings { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840032.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowTabs { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837933.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowSpaces { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191746.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowParagraphs { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192337.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowHyphens { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839598.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowHiddenText { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840568.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool WrapToWindow { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197748.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowPicturePlaceHolders { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192158.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowBookmarks { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195455.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdFieldShading FieldShading { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowAnimation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844911.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool TableGridlines { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 EnlargeFontsLessThan { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845503.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowMainTextLayer { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834537.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdSeekView SeekView { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196649.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdSpecialPane SplitSpecial { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 BrowseToWindow { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839824.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowOptionalBreaks { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197553.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool DisplayPageBoundaries { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool DisplaySmartTags { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836994.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ShowRevisionsAndComments { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844875.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ShowComments { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193645.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ShowInsertionsAndDeletions { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839560.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ShowFormatChanges { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdRevisionsView RevisionsView { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdRevisionsMode RevisionsMode { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840523.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Single RevisionsBalloonWidth { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840447.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdRevisionsBalloonWidthType RevisionsBalloonWidthType { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197146.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdRevisionsBalloonMargin RevisionsBalloonSide { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Reviewers Reviewers { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821251.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool RevisionsBalloonShowConnectingLines { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839617.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool ReadingLayout { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198229.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 ShowXMLMarkup { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840275.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 ShadeEditableRanges { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196812.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool ShowInkAnnotations { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197827.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool DisplayBackgrounds { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197998.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool ReadingLayoutActualView { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool ReadingLayoutAllowMultiplePages { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ReadingLayoutAllowEditing { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845085.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdReadingLayoutMargin ReadingLayoutTruncateMargins { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194113.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ShowMarkupAreaHighlight { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844988.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool Panning { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837472.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ShowCropMarks { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192820.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdRevisionsMode MarkupMode { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839848.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool ConflictMode { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845837.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool ShowOtherAuthors { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231107.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.RevisionsFilter RevisionsFilter { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231031.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdPageColor PageColor { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230868.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdColumnWidth ColumnWidth { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836702.aspx </remarks>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CollapseOutline(object range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836702.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CollapseOutline();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194789.aspx </remarks>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ExpandOutline(object range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194789.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ExpandOutline();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192618.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ShowAllHeadings();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836914.aspx </remarks>
		/// <param name="level">Int32 level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ShowHeading(Int32 level);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841033.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PreviousHeaderFooter();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195041.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void NextHeaderFooter();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231576.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void ExpandAllHeadings();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227347.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void CollapseAllHeadings();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[SupportByVersion("Word", 15, 16)]
		void ForceOffscreenUpdate();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[SupportByVersion("Word", 15, 16)]
		void ForceLowresUpdate();

		#endregion
	}
}
