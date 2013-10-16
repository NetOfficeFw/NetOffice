using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoFarEastLineBreakLanguageID
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1041</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 MsoFarEastLineBreakLanguageJapanese = 1041,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1042</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 MsoFarEastLineBreakLanguageKorean = 1042,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2052</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 MsoFarEastLineBreakLanguageSimplifiedChinese = 2052,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1028</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 MsoFarEastLineBreakLanguageTraditionalChinese = 1028
	}
}