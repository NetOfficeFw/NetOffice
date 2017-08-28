using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823232.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdFindMatch
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65551</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchParagraphMark = 65551,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchTabCharacter = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchCommentMark = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65599</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchAnyCharacter = 65599,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65567</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchAnyDigit = 65567,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65583</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchAnyLetter = 65583,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchCaretCharacter = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchColumnBreak = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8212</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchEmDash = 8212,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8211</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchEnDash = 8211,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65555</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchEndnoteMark = 65555,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchField = 19,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65554</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchFootnoteMark = 65554,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchGraphic = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65551</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchManualLineBreak = 65551,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65564</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchManualPageBreak = 65564,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchNonbreakingHyphen = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchNonbreakingSpace = 160,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchOptionalHyphen = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65580</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchSectionBreak = 65580,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65655</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdMatchWhiteSpace = 65655
	}
}