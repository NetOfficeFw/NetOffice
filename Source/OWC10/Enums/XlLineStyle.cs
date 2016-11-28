using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlLineStyle
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlLineStyleNone = -4142,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlContinuous = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4115</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlDash = -4115,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4118</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlDot = -4118,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlDashDot = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 xlDashDotDot = 5
	}
}