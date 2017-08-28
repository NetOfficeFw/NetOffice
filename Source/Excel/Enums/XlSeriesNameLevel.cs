using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231659.aspx </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlSeriesNameLevel
	{
		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlSeriesNameLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlSeriesNameLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlSeriesNameLevelAll = -1
	}
}