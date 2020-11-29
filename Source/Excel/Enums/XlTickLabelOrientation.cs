using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlTickLabelOrientation"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlTickLabelOrientation
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTickLabelOrientationAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTickLabelOrientationDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTickLabelOrientationHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTickLabelOrientationUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlTickLabelOrientationVertical = -4166
	}
}