using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837609.aspx </remarks>
	[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPivotTableVersionList
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		 xlPivotTableVersion2000 = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		 xlPivotTableVersion10 = 1,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		 xlPivotTableVersionCurrent = -1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlPivotTableVersion11 = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlPivotTableVersion12 = 3,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 14,15,16)]
		 xlPivotTableVersion14 = 4,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlPivotTableVersion15 = 5
	}
}