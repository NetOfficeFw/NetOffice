using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCubeFieldSubType
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeHierarchy = 1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeMeasure = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeSet = 3,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeAttribute = 4,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeCalculatedMeasure = 5,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeKPIValue = 6,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeKPIGoal = 7,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeKPIStatus = 8,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeKPITrend = 9,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlCubeKPIWeight = 10
	}
}