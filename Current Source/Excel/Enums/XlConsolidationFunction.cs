using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlConsolidationFunction
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4106</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlAverage = -4106,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4112</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCount = -4112,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4113</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlCountNums = -4113,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4136</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlMax = -4136,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4139</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlMin = -4139,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4149</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlProduct = -4149,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4155</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlStDev = -4155,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4156</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlStDevP = -4156,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4157</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlSum = -4157,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4164</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlVar = -4164,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4165</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlVarP = -4165,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlUnknown = 1000
	}
}