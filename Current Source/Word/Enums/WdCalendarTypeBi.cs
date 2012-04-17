using System;
using LateBindingApi.Core;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCalendarTypeBi
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdCalendarTypeBidi = 99,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdCalendarTypeGregorian = 100
	}
}