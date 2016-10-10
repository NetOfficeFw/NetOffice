using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 15,16
	 /// </summary>
	[SupportByVersionAttribute("Visio", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisReplaceFlags
	{
		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visReplaceShapeDefault = 0,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visReplaceShapeKeepBasic = 1,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visReplaceShapeLockText = 2,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visReplaceShapeLockShapeData = 4,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visReplaceShapeLockFormat = 8
	}
}