using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpAlertLevel
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		 ppAlertsNone = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		 ppAlertsAll = 2
	}
}