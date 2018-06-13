using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// Selection
    /// </summary>
    [SyntaxBypass]
    public interface Selection_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dataOnly">optional bool dataOnly</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_XML(object dataOnly);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_XML
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx </remarks>
        /// <param name="dataOnly">optional bool dataOnly</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_XML")]
        string XML(object dataOnly);

        #endregion
    }
  
    /// <summary>
    /// DispatchInterface Selection 
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821411.aspx </remarks>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00020975-0000-0000-C000-000000000046")]
    public interface Selection : Selection_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192754.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836670.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range FormattedText { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839485.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Start { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834869.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 End { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837859.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Font Font { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821048.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdSelectionType Type { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191739.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdStoryType StoryType { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838978.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        object Style { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845908.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Tables Tables { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837460.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Words Words { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193720.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Sentences Sentences { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196946.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Characters Characters { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197009.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Footnotes Footnotes { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841006.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Endnotes Endnotes { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823219.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Comments Comments { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195296.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Cells Cells { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836277.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Sections Sections { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840393.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Paragraphs Paragraphs { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193012.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Borders Borders { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192021.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Shading Shading { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845839.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Fields Fields { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838906.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FormFields FormFields { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838307.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Frames Frames { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193858.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ParagraphFormat ParagraphFormat { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197430.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.PageSetup PageSetup { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193356.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Bookmarks Bookmarks { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836357.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StoryLength { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838983.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdLanguageID LanguageID { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196398.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191830.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838134.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Hyperlinks Hyperlinks { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194663.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Columns Columns { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821842.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Rows Rows { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836744.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.HeaderFooter HeaderFooter { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845161.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool IsEndOfRowMark { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840519.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BookmarkID { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193388.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PreviousBookmarkID { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197434.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Find Find { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845594.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Range { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Information(NetOffice.WordApi.Enums.WdInformation type);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Information
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Information")]
        object Information(NetOffice.WordApi.Enums.WdInformation type);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837479.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdSelectionFlags Flags { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835497.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool Active { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820824.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool StartIsActive { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822970.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool IPAtEndOfLine { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821400.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool ExtendMode { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839310.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool ColumnSelectMode { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821992.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdTextOrientation Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193084.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.InlineShapes InlineShapes { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192167.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196980.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839166.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844964.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Document Document { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836759.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ShapeRange ShapeRange { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 NoProofing { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821380.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Tables TopLevelTables { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192601.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool LanguageDetected { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821699.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single FitTextWidth { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198226.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.HTMLDivisions HTMLDivisions { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SmartTags SmartTags { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191940.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ShapeRange ChildShapeRange { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191804.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        bool HasChildShapeRange { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845098.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FootnoteOptions FootnoteOptions { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192368.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.EndnoteOptions EndnoteOptions { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNodes XMLNodes { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode XMLParentNode { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837314.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Editors Editors { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string XML { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840039.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        object EnhMetaFileBits { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838161.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.OMaths OMaths { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820971.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        string WordOpenXML { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ContentControls ContentControls { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.ContentControl ParentContentControl { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845714.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Select();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192352.aspx </remarks>
        /// <param name="start">Int32 start</param>
        /// <param name="end">Int32 end</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SetRange(Int32 start, Int32 end);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx </remarks>
        /// <param name="direction">optional object direction</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Collapse(object direction);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Collapse();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845077.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertBefore(string text);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192184.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertAfter(string text);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Next(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Next();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Next(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Previous(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Previous();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Previous(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartOf(object unit, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartOf();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartOf(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndOf(object unit, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndOf();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndOf(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Move(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Move();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Move(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStart(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStart();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStart(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEnd(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEnd();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEnd(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveWhile(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveWhile(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartWhile(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartWhile(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndWhile(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndWhile(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUntil(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUntil(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartUntil(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartUntil(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndUntil(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndUntil(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192037.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Cut();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196538.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840284.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertBreak(object type);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertBreak();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        /// <param name="link">optional object link</param>
        /// <param name="attachment">optional object attachment</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName, object range, object confirmConversions, object link, object attachment);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName, object range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName, object range, object confirmConversions);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName, object range, object confirmConversions, object link);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192633.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool InStory(NetOffice.WordApi.Range range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193660.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool InRange(NetOffice.WordApi.Range range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Expand(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Expand();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837485.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertParagraph();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836408.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertParagraphAfter();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTimeOld(object dateTimeFormat, object insertAsField, object insertAsFullWidth);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTimeOld();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTimeOld(object dateTimeFormat);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTimeOld(object dateTimeFormat, object insertAsField);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        /// <param name="unicode">optional object unicode</param>
        /// <param name="bias">optional object bias</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber, object font, object unicode, object bias);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber, object font);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        /// <param name="unicode">optional object unicode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber, object font, object unicode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        /// <param name="separateNumbers">optional object separateNumbers</param>
        /// <param name="separatorString">optional object separatorString</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers, object separatorString);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        /// <param name="separateNumbers">optional object separateNumbers</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title, object titleAutoText, object position);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        /// <param name="excludeLabel">optional object excludeLabel</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title, object titleAutoText, object position, object excludeLabel);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
        /// <param name="label">object label</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title, object titleAutoText);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840576.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CopyAsPicture();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="languageID">optional object languageID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object languageID);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821863.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortAscending();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845052.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortDescending();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196258.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool IsEqual(NetOffice.WordApi.Range range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835748.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single Calculate();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        /// <param name="count">optional object count</param>
        /// <param name="name">optional object name</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what, object which, object count, object name);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
        /// <param name="what">optional object what</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what, object which);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what, object which, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836451.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839107.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="dataType">optional object dataType</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconLabel">optional object iconLabel</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName, object iconLabel);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link, object placement);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="dataType">optional object dataType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="dataType">optional object dataType</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834516.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Field PreviousField();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194299.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Field NextField();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840515.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertParagraphBefore();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx </remarks>
        /// <param name="shiftCells">optional object shiftCells</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCells(object shiftCells);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCells();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx </remarks>
        /// <param name="character">optional object character</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Extend(object character);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Extend();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840081.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Shrink();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveLeft(object unit, object count, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveLeft();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveLeft(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveLeft(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveRight(object unit, object count, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveRight();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveRight(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveRight(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUp(object unit, object count, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUp();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUp(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUp(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveDown(object unit, object count, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveDown();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveDown(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveDown(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 HomeKey(object unit, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 HomeKey();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 HomeKey(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndKey(object unit, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndKey();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndKey(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835736.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void EscapeKey();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840867.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TypeText(string text);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840230.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CopyFormat();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196637.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteFormat();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839799.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TypeParagraph();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194909.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TypeBackspace();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839790.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void NextSubdocument();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845750.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PreviousSubdocument();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836022.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectColumn();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197469.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCurrentFont();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822643.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCurrentAlignment();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191872.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCurrentSpacing();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193883.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCurrentIndent();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193718.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCurrentTabs();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840690.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCurrentColor();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839540.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CreateTextbox();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840046.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void WholeStory();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845469.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectRow();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196707.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SplitTable();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx </remarks>
        /// <param name="numRows">optional object numRows</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertRows(object numRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertRows();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838759.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertColumns();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
        /// <param name="formula">optional object formula</param>
        /// <param name="numberFormat">optional object numberFormat</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFormula(object formula, object numberFormat);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFormula();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
        /// <param name="formula">optional object formula</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFormula(object formula);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx </remarks>
        /// <param name="wrap">optional object wrap</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Revision NextRevision(object wrap);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Revision NextRevision();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx </remarks>
        /// <param name="wrap">optional object wrap</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Revision PreviousRevision(object wrap);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Revision PreviousRevision();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194535.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteAsNestedTable();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839331.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="styleName">string styleName</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.AutoTextEntry CreateAutoTextEntry(string name, string styleName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838494.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DetectLanguage();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195143.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SelectCell();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx </remarks>
        /// <param name="numRows">optional object numRows</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertRowsBelow(object numRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertRowsBelow();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844950.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertColumnsRight();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx </remarks>
        /// <param name="numRows">optional object numRows</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertRowsAbove(object numRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertRowsAbove();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821034.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void RtlRun();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839502.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void LtrRun();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845275.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void BoldRun();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845442.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ItalicRun();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836904.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void RtlPara();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834853.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void LtrPara();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        /// <param name="dateLanguage">optional object dateLanguage</param>
        /// <param name="calendarType">optional object calendarType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage, object calendarType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat, object insertAsField);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        /// <param name="dateLanguage">optional object dateLanguage</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        /// <param name="subFieldNumber">optional object subFieldNumber</param>
        /// <param name="subFieldNumber2">optional object subFieldNumber2</param>
        /// <param name="subFieldNumber3">optional object subFieldNumber3</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2, object subFieldNumber3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        /// <param name="subFieldNumber">optional object subFieldNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        /// <param name="subFieldNumber">optional object subFieldNumber</param>
        /// <param name="subFieldNumber2">optional object subFieldNumber2</param>
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        /// <param name="autoFitBehavior">optional object autoFitBehavior</param>
        /// <param name="defaultTableBehavior">optional object defaultTableBehavior</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior, object defaultTableBehavior);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        /// <param name="autoFitBehavior">optional object autoFitBehavior</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000();

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197496.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void ClearFormatting();

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196969.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PasteAppendTable();

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839633.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void ToggleCharacterCode();

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821674.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType type</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837670.aspx </remarks>
        /// <param name="linkedToExcel">bool linkedToExcel</param>
        /// <param name="wordFormatting">bool wordFormatting</param>
        /// <param name="rTF">bool rTF</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rTF);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838352.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void ShrinkDiscontiguousSelection();

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838293.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void InsertStyleSeparator();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCaptionXP(object label, object title, object titleAutoText, object position);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCaptionXP(object label);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCaptionXP(object label, object title);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertCaptionXP(object label, object title, object titleAutoText);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx </remarks>
        /// <param name="editorID">optional object editorID</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToEditableRange(object editorID);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToEditableRange();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="transform">optional object transform</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertXML(string xML, object transform);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx </remarks>
        /// <param name="xML">string xML</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertXML(string xML);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838493.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ClearParagraphStyle();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191975.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ClearCharacterAllFormatting();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841083.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ClearCharacterStyle();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838672.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ClearCharacterDirectFormatting();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        /// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        /// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        /// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        /// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196419.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ReadingModeGrowFont();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196279.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ReadingModeShrinkFont();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836876.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ClearParagraphAllFormatting();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197502.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ClearParagraphDirectFormatting();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195985.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void InsertNewPage();

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings();

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe);

        #endregion
    }
}
