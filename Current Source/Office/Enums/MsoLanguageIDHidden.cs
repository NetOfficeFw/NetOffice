using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoLanguageIDHidden
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3076</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoLanguageIDChineseHongKong = 3076,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5124</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoLanguageIDChineseMacao = 5124,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11273</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoLanguageIDEnglishTrinidad = 11273
	}
}