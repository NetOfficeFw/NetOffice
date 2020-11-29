using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlSortDataOption"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlSortDataOption
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlSortNormal = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlSortTextAsNumbers = 1
	}
}