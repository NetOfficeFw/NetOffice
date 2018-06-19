using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface TextRange 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002124A-0000-0000-C000-000000000046")]
	public interface TextRange : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string Text { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Font Font { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Length { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Start { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 End { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single BoundLeft { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single BoundHeight { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single BoundTop { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single BoundWidth { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ParagraphFormat ParagraphFormat { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object ContainingObject { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Duplicate { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Font MajorityFont { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ParagraphFormat MajorityParagraphFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Fields Fields { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Story Story { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoLanguageID LanguageID { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.DropCap DropCap { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbFontScriptType Script { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlinks Hyperlinks { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.FindReplace Find { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.InlineShapes InlineShapes { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 WordsCount { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 LinesCount { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 ParagraphsCount { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.PublisherApi.Enums.PbCollapseDirection direction</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Collapse(NetOffice.PublisherApi.Enums.PbCollapseDirection direction);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Expand(NetOffice.PublisherApi.Enums.PbTextUnit unit);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		/// <param name="size">Int32 size</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Move(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		/// <param name="size">Int32 size</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 MoveStart(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		/// <param name="size">Int32 size</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 MoveEnd(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Characters(Int32 start, object length);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Characters(Int32 start);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="newText">string newText</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertAfter(string newText);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="newText">string newText</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertBefore(string newText);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="fontName">string fontName</param>
		/// <param name="charIndex">Int32 charIndex</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertSymbol(string fontName, Int32 charIndex);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoLanguageID Language = 0</param>
		/// <param name="calendar">optional NetOffice.PublisherApi.Enums.PbCalendarType Calendar = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth, object language, object calendar);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoLanguageID Language = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth, object language);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Paragraphs(Int32 start, object length);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Paragraphs(Int32 start);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Lines(Int32 start, object length);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Lines(Int32 start);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Words(Int32 start, object length);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Words(Int32 start);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Select();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Cut();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Copy();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange Paste();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="varIndex">object varIndex</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertMailMergeField(object varIndex);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PublisherApi.Enums.PbPageNumberType Type = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertPageNumber(object type);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertPageNumber();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InsertBarcode();

		#endregion
	}
}
