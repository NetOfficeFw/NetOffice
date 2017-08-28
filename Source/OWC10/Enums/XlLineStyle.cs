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
	public enum XlLineStyle
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlLineStyleNone = -4142,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlContinuous = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4115</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlDash = -4115,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4118</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlDot = -4118,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlDashDot = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlDashDotDot = 5
	}
}