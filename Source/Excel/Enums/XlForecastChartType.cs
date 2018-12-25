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
	public enum XlForecastChartType
	{
		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastChartTypeLine = 0,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastChartTypeColumn = 1
	}
}