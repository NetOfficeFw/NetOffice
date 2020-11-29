using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlCVError"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlCVError
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2007</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrDiv0 = 2007,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2042</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrNA = 2042,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2029</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrName = 2029,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrNull = 2000,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2036</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrNum = 2036,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2023</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrRef = 2023,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2015</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlErrValue = 2015
	}
}