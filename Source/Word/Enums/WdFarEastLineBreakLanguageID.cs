using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdFarEastLineBreakLanguageID"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdFarEastLineBreakLanguageID
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1041</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdLineBreakJapanese = 1041,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1042</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdLineBreakKorean = 1042,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2052</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdLineBreakSimplifiedChinese = 2052,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1028</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdLineBreakTraditionalChinese = 1028
	}
}