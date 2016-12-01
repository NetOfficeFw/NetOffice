using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlDirection
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4121</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlDown = -4121,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlToRight = -4161,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlUp = -4162
	}
}