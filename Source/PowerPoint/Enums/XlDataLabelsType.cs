using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlDataLabelsType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlDataLabelsShowNone = -4142,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlDataLabelsShowValue = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlDataLabelsShowPercent = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlDataLabelsShowLabel = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlDataLabelsShowLabelAndPercent = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlDataLabelsShowBubbleSizes = 6
	}
}