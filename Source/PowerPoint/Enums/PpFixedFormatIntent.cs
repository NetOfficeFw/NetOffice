using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpFixedFormatIntent"/> </remarks>
	[SupportByVersion("PowerPoint", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpFixedFormatIntent
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppFixedFormatIntentScreen = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppFixedFormatIntentPrint = 2
	}
}