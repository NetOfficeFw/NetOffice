using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227360.aspx </remarks>
	[SupportByVersionAttribute("Excel", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCategoryLabelLevel
	{
		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlCategoryLabelLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlCategoryLabelLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlCategoryLabelLevelAll = -1
	}
}