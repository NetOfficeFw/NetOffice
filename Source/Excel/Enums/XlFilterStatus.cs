using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlFilterStatus
	{
		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 15)]
		 xlFilterStatusOK = 0,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 15)]
		 xlFilterStatusDateWrongOrder = 1,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 15)]
		 xlFilterStatusDateHasTime = 2,

		 /// <summary>
		 /// SupportByVersion Excel 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 15)]
		 xlFilterStatusInvalidDate = 3
	}
}