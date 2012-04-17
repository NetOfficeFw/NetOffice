using System;
using LateBindingApi.Core;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoClickState
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 msoClickStateAfterAllAnimations = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 msoClickStateBeforeAutomaticAnimations = -1
	}
}