using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 16
	 /// </summary>
	[SupportByVersion("Excel", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlForecastDataCompletion
	{
		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastDataCompletionZeros = 0,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastDataCompletionInterpolate = 1
	}
}