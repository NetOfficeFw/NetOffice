using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.xlquickanalysismode"/> </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlQuickAnalysisMode
	{
		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlLensOnly = 0,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlFormatConditions = 1,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlRecommendedCharts = 2,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlTotals = 3,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlTables = 4,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlSparklines = 5
	}
}