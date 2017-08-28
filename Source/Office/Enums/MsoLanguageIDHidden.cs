using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoLanguageIDHidden
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3076</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoLanguageIDChineseHongKong = 3076,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5124</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoLanguageIDChineseMacao = 5124,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11273</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoLanguageIDEnglishTrinidad = 11273
	}
}