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
	public enum XlOrientation
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlVertical = -4166
	}
}