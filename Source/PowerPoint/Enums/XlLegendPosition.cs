using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.XlLegendPosition"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlLegendPosition
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlLegendPositionBottom = -4107,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlLegendPositionCorner = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlLegendPositionLeft = -4131,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlLegendPositionRight = -4152,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlLegendPositionTop = -4160,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 xlLegendPositionCustom = -4161
	}
}