using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlPivotTableMissingItems"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlPivotTableMissingItems
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlMissingItemsDefault = -1,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlMissingItemsNone = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32500</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlMissingItemsMax = 32500,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlMissingItemsMax2 = 1048576
	}
}