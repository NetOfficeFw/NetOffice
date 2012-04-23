using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlWindowState
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4137</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlMaximized = -4137,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4140</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlMinimized = -4140,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4143</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlNormal = -4143
	}
}