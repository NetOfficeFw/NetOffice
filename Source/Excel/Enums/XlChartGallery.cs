using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlChartGallery"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlChartGallery
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBuiltIn = 21,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlUserDefined = 22,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlAnyGallery = 23
	}
}