using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 15,16
	 /// </summary>
	[SupportByVersion("Visio", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisSVGExportFormat
	{
		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Visio", 15, 16)]
		 visSVGIncludeVisioElements = 0,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 15, 16)]
		 visSVGExcludeVisioElements = 1
	}
}