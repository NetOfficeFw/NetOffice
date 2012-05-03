using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlDataLabelsType
	{
		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDataLabelsShowNone = -4142,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDataLabelsShowValue = 2,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDataLabelsShowPercent = 3,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDataLabelsShowLabel = 4,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDataLabelsShowLabelAndPercent = 5,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 xlDataLabelsShowBubbleSizes = 6
	}
}