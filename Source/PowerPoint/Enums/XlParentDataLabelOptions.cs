using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 16
	 /// </summary>
	[SupportByVersion("PowerPoint", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlParentDataLabelOptions
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlParentDataLabelOptionsNone = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlParentDataLabelOptionsBanner = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlParentDataLabelOptionsOverlapping = 2
	}
}