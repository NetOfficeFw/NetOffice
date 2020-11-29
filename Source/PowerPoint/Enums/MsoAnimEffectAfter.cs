using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.MsoAnimEffectAfter"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAnimEffectAfter
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectAfterFreeze = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectAfterRemove = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectAfterHold = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectAfterTransition = 4
	}
}