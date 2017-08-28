using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsEnum)]
	public enum ChartAxisPositionEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionTop = -1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionBottom = -2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionLeft = -3,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionRight = -4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionRadial = -5,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionCircular = -6,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionCategory = -7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionTimescale = -7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionValue = -8,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionSeries = -9,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-10</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionPrimary = -10,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-11</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chAxisPositionSecondary = -11
	}
}