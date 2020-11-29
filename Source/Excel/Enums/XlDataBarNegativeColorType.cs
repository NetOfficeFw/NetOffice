using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlDataBarNegativeColorType"/> </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlDataBarNegativeColorType
	{
		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlDataBarColor = 0,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlDataBarSameAsPositive = 1
	}
}