using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdNumberStyleWordBasicBiDi"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdNumberStyleWordBasicBiDi
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleBidi1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleBidi2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleBidiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleBidiLetter2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleBidiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleBidiLetter2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdPageNumberStyleBidiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdPageNumberStyleBidiLetter2 = 50
	}
}