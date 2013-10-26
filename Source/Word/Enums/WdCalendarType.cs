using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194619.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCalendarType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCalendarWestern = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCalendarArabic = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCalendarHebrew = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12)]
		 wdCalendarChina = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCalendarJapan = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCalendarThai = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCalendarKorean = 6,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCalendarSakaEra = 7,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdCalendarTranslitEnglish = 8,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdCalendarTranslitFrench = 9,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 wdCalendarTaiwan = 3,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 wdCalendarUmalqura = 13
	}
}