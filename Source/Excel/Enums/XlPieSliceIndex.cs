using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlPieSliceIndex"/> </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlPieSliceIndex
	{
		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlOuterCounterClockwisePoint = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlOuterCenterPoint = 2,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlOuterClockwisePoint = 3,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlMidClockwiseRadiusPoint = 4,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlCenterPoint = 5,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlMidCounterClockwiseRadiusPoint = 6,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlInnerClockwisePoint = 7,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlInnerCenterPoint = 8,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlInnerCounterClockwisePoint = 9
	}
}