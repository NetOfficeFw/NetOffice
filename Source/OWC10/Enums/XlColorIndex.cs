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
	public enum XlColorIndex
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4105</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlColorIndexAutomatic = -4105,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlColorIndexNone = -4142
	}
}