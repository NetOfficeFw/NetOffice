﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlListDataType"/> </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlListDataType
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeNone = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeText = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeMultiLineText = 2,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeNumber = 3,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeCurrency = 4,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeDateTime = 5,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeChoice = 6,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeChoiceMulti = 7,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeListLookup = 8,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeCheckbox = 9,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeHyperLink = 10,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeCounter = 11,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Excel", 11,12,14,15,16)]
		 xlListDataTypeMultiLineRichText = 12
	}
}