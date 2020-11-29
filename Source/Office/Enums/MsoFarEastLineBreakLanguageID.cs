using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoFarEastLineBreakLanguageID"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoFarEastLineBreakLanguageID
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1041</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 MsoFarEastLineBreakLanguageJapanese = 1041,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1042</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 MsoFarEastLineBreakLanguageKorean = 1042,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2052</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 MsoFarEastLineBreakLanguageSimplifiedChinese = 2052,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1028</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 MsoFarEastLineBreakLanguageTraditionalChinese = 1028
	}
}