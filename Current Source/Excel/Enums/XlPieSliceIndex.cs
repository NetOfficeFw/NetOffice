using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPieSliceIndex
	{
		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlOuterCounterClockwisePoint = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlOuterCenterPoint = 2,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlOuterClockwisePoint = 3,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlMidClockwiseRadiusPoint = 4,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlCenterPoint = 5,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlMidCounterClockwiseRadiusPoint = 6,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlInnerClockwisePoint = 7,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlInnerCenterPoint = 8,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlInnerCounterClockwisePoint = 9
	}
}