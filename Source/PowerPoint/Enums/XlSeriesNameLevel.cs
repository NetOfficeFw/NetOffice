using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 15
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlSeriesNameLevel
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 15
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15)]
		 xlSeriesNameLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15)]
		 xlSeriesNameLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15)]
		 xlSeriesNameLevelAll = -1
	}
}