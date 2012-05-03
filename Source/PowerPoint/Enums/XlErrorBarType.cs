using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlErrorBarType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-4114</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlErrorBarTypeCustom = -4114,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlErrorBarTypeFixedValue = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlErrorBarTypePercent = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-4155</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlErrorBarTypeStDev = -4155,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlErrorBarTypeStError = 4
	}
}