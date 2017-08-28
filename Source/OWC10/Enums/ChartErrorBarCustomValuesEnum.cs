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
	public enum ChartErrorBarCustomValuesEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chErrorBarPlusValues = 12,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chErrorBarMinusValues = 13
	}
}