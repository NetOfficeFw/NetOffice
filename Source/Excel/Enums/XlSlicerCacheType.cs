using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231153.aspx </remarks>
	[SupportByVersionAttribute("Excel", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSlicerCacheType
	{
		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlSlicer = 1,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlTimeline = 2
	}
}