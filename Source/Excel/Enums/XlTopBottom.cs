using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlTopBottom"/> </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlTopBottom
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlTop10Top = 1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlTop10Bottom = 0
	}
}