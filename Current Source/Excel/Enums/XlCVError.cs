using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCVError
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2007</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrDiv0 = 2007,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2042</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrNA = 2042,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2029</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrName = 2029,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2000</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrNull = 2000,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2036</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrNum = 2036,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2023</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrRef = 2023,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2015</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14)]
		 xlErrValue = 2015
	}
}