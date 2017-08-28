using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820815.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlListObjectSourceType
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlSrcExternal = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlSrcRange = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlSrcXml = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlSrcQuery = 3,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlSrcModel = 4
	}
}