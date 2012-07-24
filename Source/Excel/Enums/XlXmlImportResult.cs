using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlXmlImportResult
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlXmlImportSuccess = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlXmlImportElementsTruncated = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15)]
		 xlXmlImportValidationFailed = 2
	}
}