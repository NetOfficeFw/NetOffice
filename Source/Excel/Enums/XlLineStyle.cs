using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlLineStyle
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlContinuous = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4115</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlDash = -4115,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlDashDot = 4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlDashDotDot = 5,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4118</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlDot = -4118,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4119</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlDouble = -4119,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSlantDashDot = 13,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlLineStyleNone = -4142
	}
}