using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlArrangeStyle"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlArrangeStyle
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlArrangeStyleCascade = 7,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlArrangeStyleHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlArrangeStyleTiled = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlArrangeStyleVertical = -4166
	}
}