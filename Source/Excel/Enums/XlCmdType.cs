using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlCmdType"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlCmdType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCmdCube = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCmdSql = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCmdTable = 3,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCmdDefault = 4,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlCmdList = 5,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlCmdTableCollection = 6,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlCmdExcel = 7,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlCmdDAX = 8
	}
}