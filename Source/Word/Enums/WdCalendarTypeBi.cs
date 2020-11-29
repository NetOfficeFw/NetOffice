using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdCalendarTypeBi"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdCalendarTypeBi
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCalendarTypeBidi = 99,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCalendarTypeGregorian = 100
	}
}