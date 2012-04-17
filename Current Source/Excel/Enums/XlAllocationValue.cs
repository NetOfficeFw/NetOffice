using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlAllocationValue
	{
		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlAllocateValue = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 14)]
		 xlAllocateIncrement = 2
	}
}