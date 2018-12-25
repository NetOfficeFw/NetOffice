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
	public enum XlForecastAggregation
	{
		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationAverage = 1,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationCount = 2,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationCountA = 3,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationMax = 4,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationMedian = 5,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationMin = 6,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlForecastAggregationSum = 7
	}
}