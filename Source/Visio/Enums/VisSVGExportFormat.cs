using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 15,16
	 /// </summary>
	[SupportByVersionAttribute("Visio", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisSVGExportFormat
	{
		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visSVGIncludeVisioElements = 0,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visSVGExcludeVisioElements = 1
	}
}