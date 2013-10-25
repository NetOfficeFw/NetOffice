using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745840.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlLegendPosition
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLegendPositionBottom = -4107,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLegendPositionCorner = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLegendPositionLeft = -4131,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLegendPositionRight = -4152,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLegendPositionTop = -4160,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlLegendPositionCustom = -4161
	}
}