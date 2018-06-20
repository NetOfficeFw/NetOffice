using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// TextRange2
    /// </summary>
    [SyntaxBypass]
    public interface TextRange2_ : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Paragraphs(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Paragraphs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Paragraphs")]
        NetOffice.OfficeApi.TextRange2 Paragraphs(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Paragraphs(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Paragraphs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Paragraphs")]
        NetOffice.OfficeApi.TextRange2 Paragraphs(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Sentences(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Sentences
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Sentences")]
        NetOffice.OfficeApi.TextRange2 Sentences(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Sentences(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Sentences
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Sentences")]
        NetOffice.OfficeApi.TextRange2 Sentences(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Words(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Words
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Words")]
        NetOffice.OfficeApi.TextRange2 Words(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Words(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Words
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Words")]
        NetOffice.OfficeApi.TextRange2 Words(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Characters")]
        NetOffice.OfficeApi.TextRange2 Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Characters(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Characters")]
        NetOffice.OfficeApi.TextRange2 Characters(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Lines(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Lines
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Lines")]
        NetOffice.OfficeApi.TextRange2 Lines(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Lines(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Lines
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Lines")]
        NetOffice.OfficeApi.TextRange2 Lines(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Runs(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Runs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Runs")]
        NetOffice.OfficeApi.TextRange2 Runs(object start, object length);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_Runs(object start);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Runs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Runs")]
        NetOffice.OfficeApi.TextRange2 Runs(object start);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_MathZones(object start, object length);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Alias for get_MathZones
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 14, 15, 16), Redirect("get_MathZones")]
        NetOffice.OfficeApi.TextRange2 MathZones(object start, object length);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.TextRange2 get_MathZones(object start);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Alias for get_MathZones
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 14, 15, 16), Redirect("get_MathZones")]
        NetOffice.OfficeApi.TextRange2 MathZones(object start);

        #endregion
    }
    
    /// <summary>
    /// DispatchInterface TextRange2 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863528.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("000C0397-0000-0000-C000-000000000046")]
    public interface TextRange2 : TextRange2_, IEnumerableProvider<NetOffice.OfficeApi.TextRange2>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863807.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861203.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862210.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 Paragraphs { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 Sentences { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 Words { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 Characters { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 Lines { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 Runs { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862198.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.ParagraphFormat2 ParagraphFormat { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860218.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Font2 Font { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861200.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Length { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861772.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Start { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863024.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Single BoundLeft { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863847.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Single BoundTop { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863508.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Single BoundWidth { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860263.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Single BoundHeight { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861366.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoLanguageID LanguageID { get; set; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        new NetOffice.OfficeApi.TextRange2 MathZones { get; }


        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.TextRange2 this[object index] { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861091.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 TrimText();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862180.aspx </remarks>
        /// <param name="newText">optional string NewText = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertAfter(object newText);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862180.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertAfter();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865495.aspx </remarks>
        /// <param name="newText">optional string NewText = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertBefore(object newText);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865495.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertBefore();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862495.aspx </remarks>
        /// <param name="fontName">string fontName</param>
        /// <param name="charNumber">Int32 charNumber</param>
        /// <param name="unicode">optional NetOffice.OfficeApi.Enums.MsoTriState Unicode = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertSymbol(string fontName, Int32 charNumber, object unicode);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862495.aspx </remarks>
        /// <param name="fontName">string fontName</param>
        /// <param name="charNumber">Int32 charNumber</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertSymbol(string fontName, Int32 charNumber);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860564.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Select();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862117.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Cut();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863743.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862838.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863850.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Paste();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862719.aspx </remarks>
        /// <param name="format">NetOffice.OfficeApi.Enums.MsoClipboardFormat format</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 PasteSpecial(NetOffice.OfficeApi.Enums.MsoClipboardFormat format);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864574.aspx </remarks>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoTextChangeCase type</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ChangeCase(NetOffice.OfficeApi.Enums.MsoTextChangeCase type);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861212.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddPeriods();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861820.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void RemovePeriods();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        /// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after, object matchCase, object wholeWords);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Find(string findWhat);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after, object matchCase);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        /// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after, object matchCase, object wholeWords);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after, object matchCase);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865241.aspx </remarks>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        /// <param name="x2">Single x2</param>
        /// <param name="y2">Single y2</param>
        /// <param name="x3">Single x3</param>
        /// <param name="y3">Single y3</param>
        /// <param name="x4">Single x4</param>
        /// <param name="y4">Single y4</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void RotatedBounds(out Single x1, out Single y1, out Single x2, out Single y2, out Single x3, out Single y3, out Single x4, out Single y4);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861210.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void RtlRun();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861750.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void LtrRun();

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227821.aspx </remarks>
        /// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
        /// <param name="formula">optional string Formula = </param>
        /// <param name="position">optional Int32 Position = -1</param>
        [SupportByVersion("Office", 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType, object formula, object position);

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227821.aspx </remarks>
        /// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType);

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227821.aspx </remarks>
        /// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
        /// <param name="formula">optional string Formula = </param>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType, object formula);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.TextRange2>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.TextRange2> GetEnumerator();

        #endregion
    }
}
