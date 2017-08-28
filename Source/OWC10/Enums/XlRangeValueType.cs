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
	public enum XlRangeValueType
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlRangeValueDefault = 10,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlRangeValueXMLSpreadsheet = 11,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlRangeValueHTML = 1000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlRangeValueCSV = 1001
	}
}