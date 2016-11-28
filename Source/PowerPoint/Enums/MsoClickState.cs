using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743934.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoClickState
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		 msoClickStateAfterAllAnimations = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		 msoClickStateBeforeAutomaticAnimations = -1
	}
}