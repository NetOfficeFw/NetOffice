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
	public enum XlDirection
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4121</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlDown = -4121,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4159</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlToLeft = -4159,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4161</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlToRight = -4161,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4162</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlUp = -4162
	}
}