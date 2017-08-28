using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdLanguageID2000
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3076</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdChineseHongKong = 3076,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5124</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdChineseMacao = 5124,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11273</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdEnglishTrinidad = 11273
	}
}