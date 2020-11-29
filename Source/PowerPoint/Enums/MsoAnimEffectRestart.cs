using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.MsoAnimEffectRestart"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAnimEffectRestart
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectRestartAlways = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectRestartWhenOff = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimEffectRestartNever = 3
	}
}