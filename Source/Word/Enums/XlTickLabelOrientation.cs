using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlTickLabelOrientation
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlTickLabelOrientationAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlTickLabelOrientationDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlTickLabelOrientationHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlTickLabelOrientationUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlTickLabelOrientationVertical = -4166
	}
}