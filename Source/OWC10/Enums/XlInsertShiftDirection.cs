using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlInsertShiftDirection
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4121</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlShiftDown = -4121,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlShiftToRight = -4161
	}
}