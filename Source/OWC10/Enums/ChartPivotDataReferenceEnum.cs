using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum ChartPivotDataReferenceEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 chPivotColumns = -1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 chPivotRows = -2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 chPivotColAggregates = -3,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 chPivotRowAggregates = -4
	}
}