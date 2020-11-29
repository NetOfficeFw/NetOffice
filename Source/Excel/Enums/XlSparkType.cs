using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlSparkType"/> </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlSparkType
	{
		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlSparkLine = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlSparkColumn = 2,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlSparkColumnStacked100 = 3
	}
}