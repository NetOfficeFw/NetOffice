using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9
	 /// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsEnum)]
	public enum PpScrollBarStyle
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppScrollBarVertical = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppScrollBarHorizontal = 1
	}
}