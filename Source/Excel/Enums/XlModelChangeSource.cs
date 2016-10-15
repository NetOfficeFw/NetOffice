using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn301197.aspx </remarks>
	[SupportByVersionAttribute("Excel", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlModelChangeSource
	{
		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlChangeByExcel = 0,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlChangeByPowerPivotAddIn = 1
	}
}