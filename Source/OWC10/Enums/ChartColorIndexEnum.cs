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
	public enum ChartColorIndexEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chColorAutomatic = -1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chColorNone = -2
	}
}