using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoAnimTextUnitEffect
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimTextUnitEffectMixed = -1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimTextUnitEffectByParagraph = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimTextUnitEffectByCharacter = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 msoAnimTextUnitEffectByWord = 2
	}
}