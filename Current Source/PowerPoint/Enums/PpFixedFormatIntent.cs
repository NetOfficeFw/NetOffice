using System;
using LateBindingApi.Core;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpFixedFormatIntent
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppFixedFormatIntentScreen = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppFixedFormatIntentPrint = 2
	}
}