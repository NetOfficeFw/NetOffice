using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlDeleteShiftDirection
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlShiftToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlShiftUp = -4162
	}
}