using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.xlfilterstatus"/> </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlFilterStatus
	{
		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlFilterStatusOK = 0,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlFilterStatusDateWrongOrder = 1,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlFilterStatusDateHasTime = 2,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 15, 16)]
		 xlFilterStatusInvalidDate = 3
	}
}