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
	public enum ChartSpecialDataSourcesEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chDataBound = 0,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chDataLiteral = -1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chDataNone = -2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 chDataLinked = -3
	}
}