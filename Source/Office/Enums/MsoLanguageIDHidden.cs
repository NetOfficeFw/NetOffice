using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoLanguageIDHidden
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3076</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		 msoLanguageIDChineseHongKong = 3076,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5124</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		 msoLanguageIDChineseMacao = 5124,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11273</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		 msoLanguageIDEnglishTrinidad = 11273
	}
}