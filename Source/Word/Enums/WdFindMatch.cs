using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823232.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdFindMatch
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65551</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchParagraphMark = 65551,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchTabCharacter = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchCommentMark = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65599</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchAnyCharacter = 65599,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65567</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchAnyDigit = 65567,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65583</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchAnyLetter = 65583,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchCaretCharacter = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchColumnBreak = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8212</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchEmDash = 8212,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8211</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchEnDash = 8211,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65555</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchEndnoteMark = 65555,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchField = 19,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65554</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchFootnoteMark = 65554,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchGraphic = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65551</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchManualLineBreak = 65551,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65564</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchManualPageBreak = 65564,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchNonbreakingHyphen = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchNonbreakingSpace = 160,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchOptionalHyphen = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65580</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchSectionBreak = 65580,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65655</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMatchWhiteSpace = 65655
	}
}