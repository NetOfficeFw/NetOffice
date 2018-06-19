using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface DropCap 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("37B3B0AF-44B5-11D3-B65B-00C04F8EF32D")]
	public interface DropCap : ICOMObject
	{
		#region Properties

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
		string FontName { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ColorFormat FontColor { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState FontBold { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState FontItalic { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 LinesUp { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Size { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Span { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="bold">optional bool Bold = false</param>
		/// <param name="italic">optional bool Italic = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap(object linesUp, object size, object span, object fontName, object bold, object italic);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap(object linesUp);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap(object linesUp, object size);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap(object linesUp, object size, object span);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap(object linesUp, object size, object span, object fontName);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="bold">optional bool Bold = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyCustomDropCap(object linesUp, object size, object span, object fontName, object bold);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Clear();

		#endregion
	}
}
