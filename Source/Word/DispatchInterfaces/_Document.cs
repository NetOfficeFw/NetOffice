using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	public interface _Document : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196900.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822944.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845529.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840549.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196862.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object BuiltInDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195603.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object CustomDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821867.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string Path { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194977.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Bookmarks Bookmarks { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835455.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Tables Tables { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197126.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Footnotes Footnotes { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194032.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Endnotes Endnotes { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845880.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Comments Comments { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdDocumentType Type { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191749.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool AutoHyphenation { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool HyphenateCaps { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193110.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 HyphenationZone { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820862.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 ConsecutiveHyphensLimit { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822125.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Sections Sections { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836325.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Paragraphs Paragraphs { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845024.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Words Words { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194403.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Sentences Sentences { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191729.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Characters Characters { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821229.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Fields Fields { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840117.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.FormFields FormFields { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193100.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Styles Styles { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197117.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Frames Frames { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191950.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TablesOfFigures TablesOfFigures { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834524.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Variables Variables { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198370.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMerge MailMerge { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844798.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Envelope Envelope { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821285.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string FullName { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192540.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Revisions Revisions { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822932.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TablesOfContents TablesOfContents { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837912.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TablesOfAuthorities TablesOfAuthorities { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839306.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.PageSetup PageSetup { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837336.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Windows Windows { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool HasRoutingSlip { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.RoutingSlip RoutingSlip { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Routed { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838095.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TablesOfAuthoritiesCategories TablesOfAuthoritiesCategories { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Indexes Indexes { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194753.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Saved { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821853.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Content { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Window ActiveWindow { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192728.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdDocumentKind Kind { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196223.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ReadOnly { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195362.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Subdocuments Subdocuments { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840840.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool IsMasterDocument { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196079.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single DefaultTabStop { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836281.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool EmbedTrueTypeFonts { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845567.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SaveFormsData { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838914.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ReadOnlyRecommended { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844828.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SaveSubsetFonts { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840506.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool get_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type, bool value);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Compatibility
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840506.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_Compatibility")]
		bool Compatibility(NetOffice.WordApi.Enums.WdCompatibility type);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197823.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.StoryRanges StoryRanges { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821872.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.CommandBars CommandBars { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192771.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool IsSubdocument { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840755.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 SaveFormat { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836643.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdProtectionType ProtectionType { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837239.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Hyperlinks Hyperlinks { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197211.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Shapes Shapes { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839163.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ListTemplates ListTemplates { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Lists Lists { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821398.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool UpdateStylesOnOpen { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839734.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object AttachedTemplate { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844996.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.InlineShapes InlineShapes { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Shape Background { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193109.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool GrammarChecked { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845040.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SpellingChecked { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836692.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowGrammaticalErrors { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821056.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowSpellingErrors { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Versions Versions { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowSummary { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdSummaryMode SummaryViewMode { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 SummaryLength { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool PrintFractionalWidths { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196987.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool PrintPostScriptOverText { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840423.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Container { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838735.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool PrintFormsData { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198090.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ListParagraphs ListParagraphs { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192387.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string Password { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839518.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string WritePassword { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194500.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool HasPassword { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837527.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool WriteReserved { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844946.aspx </remarks>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_ActiveWritingStyle(object languageID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_ActiveWritingStyle(object languageID, string value);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ActiveWritingStyle
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844946.aspx </remarks>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_ActiveWritingStyle")]
		string ActiveWritingStyle(object languageID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193401.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool UserControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool HasMailer { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Mailer Mailer { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192400.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ProofreadingErrors GrammaticalErrors { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838118.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ProofreadingErrors SpellingErrors { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837668.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.VBIDEApi.VBProject VBProject { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840586.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool FormsDesign { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string _CodeName { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197577.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string CodeName { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821373.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SnapToGrid { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837193.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SnapToShapes { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839124.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single GridDistanceHorizontal { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195287.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single GridDistanceVertical { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839558.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single GridOriginHorizontal { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198193.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single GridOriginVertical { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821306.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 GridSpaceBetweenHorizontalLines { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821996.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 GridSpaceBetweenVerticalLines { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845752.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool GridOriginFromMargin { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836931.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool KerningByAlgorithm { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191748.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdJustificationMode JustificationMode { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845667.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdFarEastLineBreakLevel FarEastLineBreakLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844966.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string NoLineBreakBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192597.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string NoLineBreakAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838067.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool TrackRevisions { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192825.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool PrintRevisions { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ShowRevisions { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192741.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string ActiveTheme { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837037.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string ActiveThemeDisplayName { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839292.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Email Email { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196093.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Scripts Scripts { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191794.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool LanguageDetected { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838486.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID FarEastLineBreakLanguage { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194305.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Frameset Frameset { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839615.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ClickAndTypeParagraphStyle { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.HTMLProject HTMLProject { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844954.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.WebOptions WebOptions { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835467.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoEncoding OpenEncoding { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834893.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoEncoding SaveEncoding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835162.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool OptimizeForWord97 { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836069.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool VBASigned { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840465.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.MsoEnvelope MailEnvelope { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194348.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool DisableFeatures { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194604.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool DoNotEmbedSystemFonts { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193069.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.SignatureSet Signatures { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194661.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		string DefaultTargetFrame { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822985.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.HTMLDivisions HTMLDivisions { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196211.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838361.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool RemovePersonalInformation { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.SmartTags SmartTags { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool EmbedSmartTags { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool SmartTagsAsXMLProps { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835460.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoEncoding TextEncoding { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198078.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLineEndingType TextLineEnding { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845673.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.StyleSheets StyleSheets { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837042.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		object DefaultTableStyle { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194870.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		string PasswordEncryptionProvider { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195788.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		string PasswordEncryptionAlgorithm { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193119.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Int32 PasswordEncryptionKeyLength { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822966.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool PasswordEncryptionFileProperties { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836336.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool EmbedLinguisticData { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839893.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool FormattingShowFont { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839706.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool FormattingShowClear { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836749.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool FormattingShowParagraph { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193041.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool FormattingShowNumbering { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194361.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdShowFilter FormattingShowFilter { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191744.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.OfficeApi.Permission Permission { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNodes XMLNodes { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198201.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLSchemaReferences XMLSchemaReferences { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840776.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.OfficeApi.SmartDocument SmartDocument { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.OfficeApi.SharedWorkspace SharedWorkspace { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837910.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.OfficeApi.Sync Sync { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838344.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool EnforceStyle { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822185.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool AutoFormatOverride { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool XMLSaveDataOnly { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool XMLHideNamespaces { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196205.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool XMLShowAdvancedErrors { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836689.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool XMLUseXSLTWhenSaving { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838300.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		string XMLSaveThroughXSLT { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191946.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196654.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool ReadingModeLayoutFrozen { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194610.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		bool RemoveDateAndTime { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNodes XMLSchemaViolations { get; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191938.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 ReadingLayoutSizeX { get; set; }

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839167.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		Int32 ReadingLayoutSizeY { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191767.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdStyleSort StyleSortMethod { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844919.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.MetaProperties ContentTypeProperties { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197907.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool TrackMoves { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836881.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool TrackFormatting { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Dummy1 { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837488.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.OMaths OMaths { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Dummy3 { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839289.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.ServerPolicy ServerPolicy { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822382.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControls ContentControls { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839144.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.DocumentInspectors DocumentInspectors { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834552.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Bibliography Bibliography { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198209.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool LockTheme { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839340.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool LockQuickStyleSet { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821063.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string OriginalDocumentTitle { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834817.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string RevisedDocumentTitle { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193091.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.CustomXMLParts CustomXMLParts { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195284.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool FormattingShowNextLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191723.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool FormattingShowUserStyleName { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822952.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Research Research { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838930.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool Final { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821662.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdOMathBreakBin OMathBreakBin { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835681.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdOMathBreakSub OMathBreakSub { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196528.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdOMathJc OMathJc { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195080.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Single OMathLeftMargin { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192826.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Single OMathRightMargin { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195018.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Single OMathWrap { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822912.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool OMathIntSubSupLim { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192808.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool OMathNarySupSubLim { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835679.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool OMathSmallFrac { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197690.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string WordOpenXML { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840566.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.OfficeTheme DocumentTheme { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845747.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool HasVBProject { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193851.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string OMathFontName { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836379.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string EncryptionProvider { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838359.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool UseMathDefaults { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195620.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 CurrentRsid { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 DocID { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196837.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 CompatibilityMode { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837045.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.CoAuthoring CoAuthoring { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231858.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Broadcast Broadcast { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228844.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		bool ChartDataPointTrack { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230857.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		bool IsInAutosave { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close(object saveChanges, object originalFormat, object routeDocument);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close(object saveChanges);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Close(object saveChanges, object originalFormat);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821326.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Repaginate();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822617.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FitToPages();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841098.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ManualHyphenation();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845112.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Select();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845755.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DataForm();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Route();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821625.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821630.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SendMail();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="end">optional object end</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Range(object start, object end);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Range();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx </remarks>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Range(object start);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823210.aspx </remarks>
		/// <param name="which">NetOffice.WordApi.Enums.WdAutoMacros which</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RunAutoMacro(NetOffice.WordApi.Enums.WdAutoMacros which);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822131.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195898.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintPreview();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range GoTo(object what, object which, object count, object name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range GoTo();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range GoTo(object what);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range GoTo(object what, object which);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range GoTo(object what, object which, object count);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840796.aspx </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Undo(object times);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840796.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Undo();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845577.aspx </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Redo(object times);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845577.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Redo();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840638.aspx </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		/// <param name="includeFootnotesAndEndnotes">optional object includeFootnotesAndEndnotes</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic, object includeFootnotesAndEndnotes);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840638.aspx </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845133.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MakeCompatibilityDefault();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		/// <param name="useIRM">optional object useIRM</param>
		/// <param name="enforceStyleLock">optional object enforceStyleLock</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Protect(NetOffice.WordApi.Enums.WdProtectionType type);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		/// <param name="useIRM">optional object useIRM</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845016.aspx </remarks>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Unprotect(object password);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845016.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Unprotect();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption option</param>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption option</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx </remarks>
		/// <param name="letterContent">optional object letterContent</param>
		/// <param name="wizardMode">optional object wizardMode</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RunLetterWizard(object letterContent, object wizardMode);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RunLetterWizard();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx </remarks>
		/// <param name="letterContent">optional object letterContent</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RunLetterWizard(object letterContent);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836106.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent GetLetterContent();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822930.aspx </remarks>
		/// <param name="letterContent">object letterContent</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SetLetterContent(object letterContent);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840260.aspx </remarks>
		/// <param name="template">string template</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CopyStylesFromTemplate(string template);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840983.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void UpdateStyles();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834835.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckGrammar();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		/// <param name="headerInfo">optional object headerInfo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address, object subAddress);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address, object subAddress, object newWindow);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839781.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AddToFavorites();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195614.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Reload();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		/// <param name="mode">optional object mode</param>
		/// <param name="updateProperties">optional object updateProperties</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range AutoSummarize(object length, object mode, object updateProperties);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range AutoSummarize();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range AutoSummarize(object length);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		/// <param name="mode">optional object mode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range AutoSummarize(object length, object mode);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193060.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveNumbers(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193060.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveNumbers();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838874.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ConvertNumbersToText(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838874.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ConvertNumbersToText();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems(object numberType, object level);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192151.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Post();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195394.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ToggleFormsDesign();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Compare(string name);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object removePersonalInformation</param>
		/// <param name="removeDateAndTime">optional object removeDateAndTime</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation, object removeDateAndTime);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Compare(string name, object authorName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Compare(string name, object authorName, object compareTarget);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Compare(string name, object authorName, object compareTarget, object detectFormatChanges);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object removePersonalInformation</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void UpdateSummaryProperties();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193699.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		object GetCrossReferenceItems(object referenceType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193992.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837880.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ViewCode();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834519.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ViewPropertyBrowser();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ForwardMailer();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Reply();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ReplyAll();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="priority">optional object priority</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SendMailer(object fileFormat, object priority);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SendMailer();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SendMailer(object fileFormat);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195616.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void UndoClear();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192417.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PresentIt();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838927.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subject">optional object subject</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SendFax(string address, object subject);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838927.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SendFax(string address);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Merge(string fileName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object useFormattingFrom</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Merge(string fileName, object mergeTarget);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Merge(string fileName, object mergeTarget, object detectFormatChanges);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object useFormattingFrom</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822702.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ClosePrintPreview();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834920.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void CheckConsistency();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		/// <param name="senderCode">optional object senderCode</param>
		/// <param name="senderGender">optional object senderGender</param>
		/// <param name="senderReference">optional object senderReference</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender, object senderReference);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		/// <param name="senderCode">optional object senderCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		/// <param name="senderCode">optional object senderCode</param>
		/// <param name="senderGender">optional object senderGender</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193342.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AcceptAllRevisions();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838536.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RejectAllRevisions();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197127.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DetectLanguage();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835740.aspx </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyTheme(string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839088.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveTheme();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835177.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void WebPagePreview();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195768.aspx </remarks>
		/// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding encoding</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void sblt(string s);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Compare2000(string name);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Merge2000(string fileName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838511.aspx </remarks>
		/// <param name="codePageOrigin">Int32 codePageOrigin</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ConvertVietDoc(Int32 codePageOrigin);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void CheckIn(object saveChanges, object comments, object makePublic);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void CheckIn();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void CheckIn(object saveChanges);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void CheckIn(object saveChanges, object comments);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198206.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool CanCheckin();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		/// <param name="includeAttachment">optional object includeAttachment</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SendForReview(object recipients, object subject, object showMessage, object includeAttachment);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SendForReview();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SendForReview(object recipients);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SendForReview(object recipients, object subject);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SendForReview(object recipients, object subject, object showMessage);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836324.aspx </remarks>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ReplyWithChanges(object showMessage);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836324.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ReplyWithChanges();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837660.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void EndReview();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195460.aspx </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">optional object passwordEncryptionFileProperties</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, object passwordEncryptionFileProperties);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195460.aspx </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void RecheckSmartTags();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void RemoveSmartTags();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198118.aspx </remarks>
		/// <param name="style">object style</param>
		/// <param name="setInTemplate">bool setInTemplate</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SetDefaultTableStyle(object style, bool setInTemplate);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822910.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DeleteAllComments();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837501.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void AcceptAllRevisionsShown();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822533.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void RejectAllRevisionsShown();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836620.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DeleteAllCommentsShown();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821137.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ResetFormFields();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void CheckNewSmartTags();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare2002(string name);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare2002(string name, object authorName);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare2002(string name, object authorName, object compareTarget);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void SendFaxOverInternet(object recipients, object subject, object showMessage);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void SendFaxOverInternet();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void SendFaxOverInternet(object recipients);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void SendFaxOverInternet(object recipients, object subject);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196274.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="dataOnly">optional bool DataOnly = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void TransformDocument(string path, object dataOnly);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196274.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void TransformDocument(string path);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195660.aspx </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void SelectAllEditableRanges(object editorID);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195660.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void SelectAllEditableRanges();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844883.aspx </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DeleteAllEditableRanges(object editorID);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844883.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DeleteAllEditableRanges();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838947.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DeleteAllInkAnnotations();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="richFormat">bool richFormat</param>
		/// <param name="url">string url</param>
		/// <param name="title">string title</param>
		/// <param name="description">string description</param>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void AddDocumentWorkspaceHeader(bool richFormat, string url, string title, string description, string iD);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void RemoveDocumentWorkspaceHeader(string iD);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845389.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void RemoveLockedStyles();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNode SelectSingleNode(string xPath);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNodes SelectNodes(string xPath);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197270.aspx </remarks>
		/// <param name="removeDocInfoType">NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void RemoveDocumentInformation(NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void CheckInWithVersion();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges, object comments);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges, object comments, object makePublic);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		void Dummy2();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845518.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void LockServerFile();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198071.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845242.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		void Dummy4();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="skipIfAbsent">bool skipIfAbsent</param>
		/// <param name="url">string url</param>
		/// <param name="title">string title</param>
		/// <param name="description">string description</param>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		void AddMeetingWorkspaceHeader(bool skipIfAbsent, string url, string title, string description, string iD);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198291.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void SaveAsQuickStyleSet(string fileName);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyQuickStyleSet(string name);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840910.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyDocumentTheme(string fileName);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838276.aspx </remarks>
		/// <param name="node">NetOffice.OfficeApi.CustomXMLNode node</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControls SelectLinkedControls(NetOffice.OfficeApi.CustomXMLNode node);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198010.aspx </remarks>
		/// <param name="stream">optional NetOffice.OfficeApi.CustomXMLPart Stream = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControls SelectUnlinkedControls(object stream);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198010.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControls SelectUnlinkedControls();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822990.aspx </remarks>
		/// <param name="title">string title</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControls SelectContentControlsByTitle(string title);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196504.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void FreezeLayout();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		void UnfreezeLayout();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194276.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void DowngradeDocument();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835714.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void Convert();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839693.aspx </remarks>
		/// <param name="tag">string tag</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControls SelectContentControlsByTag(string tag);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838360.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		void ConvertAutoHyphens();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821672.aspx </remarks>
		/// <param name="style">object style</param>
		[SupportByVersion("Word", 14,15,16)]
		void ApplyQuickStyleSet2(object style);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		/// <param name="compatibilityMode">optional object compatibilityMode</param>
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840359.aspx </remarks>
		/// <param name="mode">Int32 mode</param>
		[SupportByVersion("Word", 14,15,16)]
		void SetCompatibilityMode(Int32 mode);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231927.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 ReturnToLastReadPosition();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		/// <param name="compatibilityMode">optional object compatibilityMode</param>
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding);

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks);

		#endregion
	}
}
