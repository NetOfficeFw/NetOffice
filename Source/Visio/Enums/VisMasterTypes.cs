using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisMasterTypes
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeMaster = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeFillPattern = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeLinePattern = 3,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeLineEnd = 4,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeDataGraphic = 5,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeThemeColors = 6,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visTypeThemeEffects = 7
	}
}