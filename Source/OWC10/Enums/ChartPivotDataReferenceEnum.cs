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
	public enum ChartPivotDataReferenceEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chPivotColumns = -1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chPivotRows = -2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chPivotColAggregates = -3,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chPivotRowAggregates = -4
	}
}