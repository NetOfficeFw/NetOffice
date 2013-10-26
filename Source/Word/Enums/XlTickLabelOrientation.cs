using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837942.aspx </remarks>
	[SupportByVersionAttribute("Word", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlTickLabelOrientation
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickLabelOrientationAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickLabelOrientationDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickLabelOrientationHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickLabelOrientationUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickLabelOrientationVertical = -4166
	}
}