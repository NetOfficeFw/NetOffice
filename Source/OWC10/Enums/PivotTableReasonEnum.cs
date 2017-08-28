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
	public enum PivotTableReasonEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plPivotTableReasonTotalAdded = 0,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plPivotTableReasonTotalDeleted = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plPivotTableReasonFieldSetAdded = 2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plPivotTableReasonFieldAdded = 3
	}
}