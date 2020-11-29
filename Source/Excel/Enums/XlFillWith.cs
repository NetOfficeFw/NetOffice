using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlFillWith"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlFillWith
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4104</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlFillWithAll = -4104,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlFillWithContents = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4122</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlFillWithFormats = -4122
	}
}