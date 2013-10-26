using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193724.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdFarEastLineBreakLanguageID
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1041</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLineBreakJapanese = 1041,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1042</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLineBreakKorean = 1042,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2052</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLineBreakSimplifiedChinese = 2052,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1028</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLineBreakTraditionalChinese = 1028
	}
}