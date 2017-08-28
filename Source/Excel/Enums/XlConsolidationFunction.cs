using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837374.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlConsolidationFunction
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4106</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlAverage = -4106,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4112</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCount = -4112,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4113</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCountNums = -4113,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4136</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlMax = -4136,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4139</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlMin = -4139,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4149</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlProduct = -4149,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4155</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStDev = -4155,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4156</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStDevP = -4156,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4157</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlSum = -4157,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4164</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVar = -4164,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4165</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlVarP = -4165,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlUnknown = 1000,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlDistinctCount = 11
	}
}