using System;
using LateBindingApi.Core;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlLegendPosition
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlLegendPositionBottom = -4107,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlLegendPositionCorner = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlLegendPositionLeft = -4131,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlLegendPositionRight = -4152,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlLegendPositionTop = -4160,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlLegendPositionCustom = -4161
	}
}