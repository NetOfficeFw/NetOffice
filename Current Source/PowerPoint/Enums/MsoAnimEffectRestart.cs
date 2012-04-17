using System;
using LateBindingApi.Core;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoAnimEffectRestart
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimEffectRestartAlways = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimEffectRestartWhenOff = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimEffectRestartNever = 3
	}
}