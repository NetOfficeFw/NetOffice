using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlScaleType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4132</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlScaleLinear = -4132,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4133</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlScaleLogarithmic = -4133
	}
}