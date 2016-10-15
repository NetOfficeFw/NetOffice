using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836192.aspx </remarks>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPivotFieldCalculation
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlDifferenceFrom = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlIndex = 9,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4143</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlNoAdditionalCalculation = -4143,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlPercentDifferenceFrom = 4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlPercentOf = 3,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlPercentOfColumn = 7,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlPercentOfRow = 6,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlPercentOfTotal = 8,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		 xlRunningTotal = 5,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlPercentOfParentRow = 10,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlPercentOfParentColumn = 11,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlPercentOfParent = 12,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlPercentRunningTotal = 13,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlRankAscending = 14,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlRankDecending = 15
	}
}