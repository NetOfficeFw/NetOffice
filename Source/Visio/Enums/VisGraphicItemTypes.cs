using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisGraphicItemTypes
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visTypeIconSet = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visTypeTextCallout = 3,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visTypeDataBar = 4,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visTypeColorByValue = 5,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visTypeHeading = 6
	}
}