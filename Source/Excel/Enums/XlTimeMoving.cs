using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15
	 /// </summary>
	[SupportByVersion("Excel", 15)]
	[EntityType(EntityType.IsEnum)]
	public enum XlTimeMoving
	{
		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeNotMoving = 0,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeMovingYearly = 1,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeMovingQuarterly = 2,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeMovingMonthly = 3,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeMovingWeekly = 4,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeMovingDaily = 5,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 15)]
		 xlTimeMovingYTD = 6
	}
}