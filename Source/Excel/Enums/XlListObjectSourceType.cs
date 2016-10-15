using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820815.aspx </remarks>
	[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlListObjectSourceType
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		 xlSrcExternal = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		 xlSrcRange = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		 xlSrcXml = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlSrcQuery = 3,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlSrcModel = 4
	}
}