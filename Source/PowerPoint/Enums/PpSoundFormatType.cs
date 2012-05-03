using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpSoundFormatType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSoundFormatMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSoundFormatNone = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSoundFormatWAV = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSoundFormatMIDI = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSoundFormatCDAudio = 3
	}
}