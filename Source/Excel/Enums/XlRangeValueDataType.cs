using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlRangeValueDataType"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlRangeValueDataType
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlRangeValueDefault = 10,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlRangeValueXMLSpreadsheet = 11,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlRangeValueMSPersistXML = 12
	}
}