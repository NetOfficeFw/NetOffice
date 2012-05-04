using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisRasterExportSizeUnits
	{
		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterPixel = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterCm = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterInch = 2
	}
}