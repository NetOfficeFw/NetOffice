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
	public enum XlDeleteShiftDirection
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlShiftToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlShiftUp = -4162
	}
}