using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlLegendPosition
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLegendPositionBottom = -4107,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLegendPositionCorner = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLegendPositionLeft = -4131,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLegendPositionRight = -4152,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLegendPositionTop = -4160,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlLegendPositionCustom = -4161
	}
}