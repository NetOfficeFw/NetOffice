using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlDeleteShiftDirection
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlShiftToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlShiftUp = -4162
	}
}