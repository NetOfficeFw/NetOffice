using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisThemeTypes
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visThemeTypeColor = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visThemeTypeEffect = 2,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visThemeTypeIndex = 0,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visThemeTypeConnector = 3,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visThemeTypeFont = 4
	}
}