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
	public enum XlVAlign
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlVAlignBottom = -4107,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlVAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersion("OWC10", 1)]
		 xlVAlignTop = -4160
	}
}