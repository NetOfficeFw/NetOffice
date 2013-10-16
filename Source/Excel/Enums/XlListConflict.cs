using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlListConflict
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlListConflictDialog = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlListConflictRetryAllConflicts = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlListConflictDiscardAllConflicts = 2,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlListConflictError = 3
	}
}