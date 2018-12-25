using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 16
	 /// </summary>
	[SupportByVersion("PowerPoint", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlBinsType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlBinsTypeAutomatic = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlBinsTypeCategorical = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlBinsTypeManual = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlBinsTypeBinSize = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 16)]
		 xlBinsTypeBinCount = 4
	}
}