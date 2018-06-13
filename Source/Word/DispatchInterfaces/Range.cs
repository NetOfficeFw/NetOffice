using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
	/// Range
	/// </summary>
	[SyntaxBypass]
    public interface Range_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dataOnly">optional bool dataOnly</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_XML(object dataOnly);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_XML
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx </remarks>
        /// <param name="dataOnly">optional bool dataOnly</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_XML")]
        string XML(object dataOnly);

        #endregion
    }
    
    /// <summary>
    /// DispatchInterface Range 
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845882.aspx </remarks>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002095E-0000-0000-C000-000000000046")]
    public interface Range : Range_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195101.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192541.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range FormattedText { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836102.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Start { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840998.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 End { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821026.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Font Font { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837543.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Duplicate { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837652.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdStoryType StoryType { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191956.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Tables Tables { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836346.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Words Words { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840991.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Sentences Sentences { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845462.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Characters Characters { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196597.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Footnotes Footnotes { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193114.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Endnotes Endnotes { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192150.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Comments Comments { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836072.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Cells Cells { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834837.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Sections Sections { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837006.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Paragraphs Paragraphs { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835448.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Borders Borders { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822980.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Shading Shading { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839529.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.TextRetrievalMode TextRetrievalMode { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845620.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Fields Fields { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834816.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FormFields FormFields { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837877.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Frames Frames { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834843.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ParagraphFormat ParagraphFormat { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195640.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ListFormat ListFormat { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195181.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Bookmarks Bookmarks { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196242.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839336.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822923.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844991.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Bold { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821583.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Italic { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821959.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdUnderline Underline { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198151.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdEmphasisMark EmphasisMark { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844978.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool DisableCharacterSpaceGrid { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838481.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Revisions Revisions { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836418.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        object Style { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845486.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StoryLength { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839161.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdLanguageID LanguageID { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837028.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SynonymInfo SynonymInfo { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838128.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Hyperlinks Hyperlinks { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838758.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ListParagraphs ListParagraphs { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837692.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Subdocuments Subdocuments { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840317.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool GrammarChecked { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196502.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool SpellingChecked { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841064.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdColorIndex HighlightColorIndex { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197474.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Columns Columns { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840908.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Rows Rows { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 CanEdit { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 CanPaste { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845343.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool IsEndOfRowMark { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845646.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BookmarkID { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191844.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PreviousBookmarkID { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195912.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Find Find { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192629.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.PageSetup PageSetup { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837242.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ShapeRange ShapeRange { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834838.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdCharacterCase Case { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834587.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Information(NetOffice.WordApi.Enums.WdInformation type);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Information
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834587.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Information")]
        object Information(NetOffice.WordApi.Enums.WdInformation type);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837707.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192406.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ProofreadingErrors GrammaticalErrors { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195285.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ProofreadingErrors SpellingErrors { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195776.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdTextOrientation Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195321.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.InlineShapes InlineShapes { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193730.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range NextStoryRange { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193321.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844803.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839349.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool LanguageDetected { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197181.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single FitTextWidth { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191976.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdHorizontalInVerticalType HorizontalInVertical { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845231.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdTwoLinesInOneType TwoLinesInOne { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195015.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CombineCharacters { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844920.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 NoProofing { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194640.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Tables TopLevelTables { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192353.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Scripts Scripts { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822135.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdCharacterWidth CharacterWidth { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840112.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdKana Kana { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821869.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BoldBi { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197717.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 ItalicBi { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196542.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string ID { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194856.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820977.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        bool ShowAll { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194311.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Document Document { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195199.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FootnoteOptions FootnoteOptions { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195039.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840972.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Editors Editors { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string XML { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192034.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        object EnhMetaFileBits { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822393.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.OMaths OMaths { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192339.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        object CharacterStyle { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196075.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        object ParagraphStyle { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196585.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        object ListStyle { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841045.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        object TableStyle { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839822.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.ContentControls ContentControls { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837448.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        string WordOpenXML { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839629.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.ContentControl ParentContentControl { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845600.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.CoAuthLocks Locks { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196284.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.CoAuthUpdates Updates { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823246.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.Conflicts Conflicts { get; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231893.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        Int32 TextVisibleOnScreen { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820813.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Select();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823262.aspx </remarks>
        /// <param name="start">Int32 start</param>
        /// <param name="end">Int32 end</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SetRange(Int32 start, Int32 end);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840825.aspx </remarks>
        /// <param name="direction">optional object direction</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Collapse(object direction);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840825.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Collapse();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836272.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertBefore(string text);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192427.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertAfter(string text);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Next(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Next();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Next(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Previous(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Previous();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Previous(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartOf(object unit, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartOf();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartOf(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndOf(object unit, object extend);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndOf();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 EndOf(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Move(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Move();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Move(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStart(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStart();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStart(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEnd(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEnd();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEnd(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192586.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveWhile(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192586.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveWhile(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838537.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartWhile(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838537.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartWhile(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835396.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndWhile(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835396.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndWhile(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840312.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUntil(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840312.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveUntil(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192403.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartUntil(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192403.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveStartUntil(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197156.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndUntil(object cset, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197156.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 MoveEndUntil(object cset);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195686.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Cut();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837718.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845105.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Paste();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835132.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertBreak(object type);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835132.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertBreak();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName, object range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertFile(string fileName, object range, object confirmConversions);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197125.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool InStory(NetOffice.WordApi.Range range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822960.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool InRange(NetOffice.WordApi.Range range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete(object unit, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837449.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void WholeStory();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838477.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Expand(object unit);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838477.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Expand();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196197.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertParagraph();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822546.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        /// <param name="unicode">optional object unicode</param>
        /// <param name="bias">optional object bias</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber, object font, object unicode, object bias);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber, object font);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        /// <param name="unicode">optional object unicode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertSymbol(Int32 characterNumber, object font, object unicode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title, object titleAutoText, object position);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertCaption(object label, object title, object titleAutoText);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836633.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193013.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortAscending();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844858.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SortDescending();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838323.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool IsEqual(NetOffice.WordApi.Range range);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821015.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single Calculate();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        /// <param name="count">optional object count</param>
        /// <param name="name">optional object name</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what, object which, object count, object name);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what, object which);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoTo(object what, object which, object count);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844826.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836673.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteSpecial(object iconIndex, object link, object placement);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835691.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void LookupNameProperties();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196924.aspx </remarks>
        /// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192827.aspx </remarks>
        /// <param name="direction">Int32 direction</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Relocate(Int32 direction);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839497.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSynonyms();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">string edition</param>
        /// <param name="format">optional object format</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SubscribeTo(string edition, object format);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">string edition</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SubscribeTo(string edition);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsRTF">optional object containsRTF</param>
        /// <param name="containsText">optional object containsText</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CreatePublisher(object edition, object containsPICT, object containsRTF, object containsText);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CreatePublisher();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CreatePublisher(object edition);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CreatePublisher(object edition, object containsPICT);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsRTF">optional object containsRTF</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CreatePublisher(object edition, object containsPICT, object containsRTF);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838952.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertAutoText();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="includeFields">optional object includeFields</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to, object includeFields);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845283.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void AutoFormat();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193931.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckGrammar();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        /// <param name="customDictionary10">optional object customDictionary10</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        /// <param name="customDictionary10">optional object customDictionary10</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821256.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertParagraphBefore();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195326.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void NextSubdocument();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195945.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PreviousSubdocument();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        /// <param name="checkHangulEnding">optional object checkHangulEnding</param>
        /// <param name="enableRecentOrdering">optional object enableRecentOrdering</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering, object customDictionary);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ConvertHangulAndHanja();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ConvertHangulAndHanja(object conversionsMode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ConvertHangulAndHanja(object conversionsMode, object fastConversion);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        /// <param name="checkHangulEnding">optional object checkHangulEnding</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        /// <param name="checkHangulEnding">optional object checkHangulEnding</param>
        /// <param name="enableRecentOrdering">optional object enableRecentOrdering</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822962.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PasteAsNestedTable();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
        /// <param name="style">object style</param>
        /// <param name="symbol">optional object symbol</param>
        /// <param name="enclosedText">optional object enclosedText</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ModifyEnclosure(object style, object symbol, object enclosedText);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
        /// <param name="style">object style</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ModifyEnclosure(object style);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
        /// <param name="style">object style</param>
        /// <param name="symbol">optional object symbol</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ModifyEnclosure(object style, object symbol);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        /// <param name="raise">optional Int32 Raise = 0</param>
        /// <param name="fontSize">optional Int32 FontSize = 0</param>
        /// <param name="fontName">optional string FontName = </param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PhoneticGuide(string text, object alignment, object raise, object fontSize, object fontName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PhoneticGuide(string text);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PhoneticGuide(string text, object alignment);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        /// <param name="raise">optional Int32 Raise = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PhoneticGuide(string text, object alignment, object raise);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        /// <param name="raise">optional Int32 Raise = 0</param>
        /// <param name="fontSize">optional Int32 FontSize = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PhoneticGuide(string text, object alignment, object raise, object fontSize);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat, object insertAsField);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Sort(object excludeHeader, object fieldNumber, object sortFieldType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195289.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DetectLanguage();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        /// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
        /// <param name="commonTerms">optional bool CommonTerms = false</param>
        /// <param name="useVariants">optional bool UseVariants = false</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TCSCConverter(object wdTCSCConverterDirection, object commonTerms, object useVariants);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TCSCConverter();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        /// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TCSCConverter(object wdTCSCConverterDirection);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        /// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
        /// <param name="commonTerms">optional bool CommonTerms = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void TCSCConverter(object wdTCSCConverterDirection, object commonTerms);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193749.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType type</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193063.aspx </remarks>
        /// <param name="linkedToExcel">bool linkedToExcel</param>
        /// <param name="wordFormatting">bool wordFormatting</param>
        /// <param name="rTF">bool rTF</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rTF);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839173.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PasteAppendTable();

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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195065.aspx </remarks>
        /// <param name="editorID">optional object editorID</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToEditableRange(object editorID);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195065.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range GoToEditableRange();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839129.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="transform">optional object transform</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertXML(string xML, object transform);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839129.aspx </remarks>
        /// <param name="xML">string xML</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void InsertXML(string xML);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822335.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="format">NetOffice.WordApi.Enums.WdSaveFormat format</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportFragment(string fileName, NetOffice.WordApi.Enums.WdSaveFormat format);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821878.aspx </remarks>
        /// <param name="level">Int16 level</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void SetListLevel(Int16 level);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191966.aspx </remarks>
        /// <param name="alignment">Int32 alignment</param>
        /// <param name="relativeTo">optional Int32 RelativeTo = 0</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void InsertAlignmentTab(Int32 alignment, object relativeTo);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191966.aspx </remarks>
        /// <param name="alignment">Int32 alignment</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void InsertAlignmentTab(Int32 alignment);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839096.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="matchDestination">optional bool MatchDestination = false</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ImportFragment(string fileName, object matchDestination);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839096.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ImportFragment(string fileName);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings();

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive);

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
