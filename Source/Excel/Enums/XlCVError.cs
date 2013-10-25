using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821947.aspx </remarks>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCVError
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2007</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrDiv0 = 2007,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2042</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrNA = 2042,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2029</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrName = 2029,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrNull = 2000,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2036</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrNum = 2036,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2023</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrRef = 2023,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2015</remarks>
		 [SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		 xlErrValue = 2015
	}
}