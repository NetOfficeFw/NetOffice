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
	public enum PivotFieldFilterFunctionEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionNone = 0,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionTopCount = 3,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionBottomCount = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionTopPercent = 5,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionBottomPercent = 6,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionTopSum = 7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plFilterFunctionBottomSum = 8
	}
}