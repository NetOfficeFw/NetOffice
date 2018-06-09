using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ISpeech 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00024466-0001-0000-C000-000000000046")]
	public interface ISpeech : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlSpeakDirection Direction { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool SpeakCellOnEnter { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		/// <param name="purge">optional object purge</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 Speak(string text, object speakAsync, object speakXML, object purge);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 Speak(string text);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 Speak(string text, object speakAsync);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		Int32 Speak(string text, object speakAsync, object speakXML);

		#endregion
	}
}
