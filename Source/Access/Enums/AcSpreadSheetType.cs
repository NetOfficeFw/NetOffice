﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcSpreadSheetType"/> </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcSpreadSheetType
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel3 = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeLotusWK1 = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeLotusWK3 = 3,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeLotusWJ2 = 4,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel5 = 5,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel7 = 5,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel4 = 6,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeLotusWK4 = 7,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel97 = 8,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel8 = 8,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acSpreadsheetTypeExcel9 = 8,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acSpreadsheetTypeExcel12 = 9,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acSpreadsheetTypeExcel12Xml = 10
	}
}