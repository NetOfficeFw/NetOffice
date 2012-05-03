using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlChartPictureType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlStackScale = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlStack = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlStretch = 1
	}
}