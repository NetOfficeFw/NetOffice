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
	public enum ChartLabelOrientationEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chLabelOrientationAutomatic = 1000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chLabelOrientationHorizontal = 0,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chLabelOrientationUpward = 90,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-90</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chLabelOrientationDownward = -90
	}
}