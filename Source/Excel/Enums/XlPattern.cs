﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlPattern"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlPattern
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternChecker = 9,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternCrissCross = 16,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4121</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternDown = -4121,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternGray16 = 17,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4124</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternGray25 = -4124,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4125</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternGray50 = -4125,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4126</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternGray75 = -4126,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternGray8 = 18,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternGrid = 15,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternLightDown = 13,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternLightHorizontal = 11,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternLightUp = 14,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternLightVertical = 12,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternNone = -4142,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternSemiGray75 = 10,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternSolid = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternUp = -4162,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPatternVertical = -4166,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4000</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlPatternLinearGradient = 4000,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4001</remarks>
		 [SupportByVersion("Excel", 12,14,15,16)]
		 xlPatternRectangularGradient = 4001
	}
}