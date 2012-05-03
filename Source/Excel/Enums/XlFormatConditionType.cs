using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlFormatConditionType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCellValue = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlExpression = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlColorScale = 3,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlDatabar = 4,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlTop10 = 5,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlIconSets = 6,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlUniqueValues = 8,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlTextString = 9,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlBlanksCondition = 10,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlTimePeriod = 11,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlAboveAverageCondition = 12,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlNoBlanksCondition = 13,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlErrorsCondition = 16,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlNoErrorsCondition = 17
	}
}