using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlPivotTableVersionList
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlPivotTableVersion2000 = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlPivotTableVersion10 = 1,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Excel", 10,11,12,14)]
		 xlPivotTableVersionCurrent = -1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlPivotTableVersion11 = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14)]
		 xlPivotTableVersion12 = 3,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlPivotTableVersion14 = 4
	}
}