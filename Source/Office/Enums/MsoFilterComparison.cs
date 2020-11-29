﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoFilterComparison"/> </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoFilterComparison
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonEqual = 0,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonNotEqual = 1,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonLessThan = 2,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonGreaterThan = 3,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonLessThanEqual = 4,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonGreaterThanEqual = 5,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonIsBlank = 6,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonIsNotBlank = 7,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonContains = 8,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFilterComparisonNotContains = 9
	}
}