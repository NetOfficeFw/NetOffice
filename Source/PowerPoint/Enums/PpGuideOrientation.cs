using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.ppguideorientation"/> </remarks>
	[SupportByVersion("PowerPoint", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpGuideOrientation
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppHorizontalGuide = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 15, 16)]
		 ppVerticalGuide = 2
	}
}