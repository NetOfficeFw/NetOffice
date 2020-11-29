using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpSoundFormatType"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpSoundFormatType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSoundFormatMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSoundFormatNone = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSoundFormatWAV = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSoundFormatMIDI = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSoundFormatCDAudio = 3
	}
}