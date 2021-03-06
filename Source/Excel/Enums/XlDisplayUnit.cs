﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlDisplayUnit"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlDisplayUnit
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlHundreds = -2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlThousands = -3,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTenThousands = -4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlHundredThousands = -5,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlMillions = -6,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTenMillions = -7,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlHundredMillions = -8,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlThousandMillions = -9,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-10</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlMillionMillions = -10
	}
}