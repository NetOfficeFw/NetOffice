using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlActionType"/> </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlActionType
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlActionTypeUrl = 1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlActionTypeRowset = 16,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlActionTypeReport = 128,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlActionTypeDrillthrough = 256
	}
}