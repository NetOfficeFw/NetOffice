using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 16
	 /// </summary>
	[SupportByVersion("MSProject", 11,12,14,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjPERTBoxStyle
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjPlainBox = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjThickBox = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjShadowBox = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjDottedBox = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjFrame = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjGrayFrame = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjPatternFrame = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjMarquee = 7
	}
}