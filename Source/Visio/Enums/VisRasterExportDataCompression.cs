using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisRasterExportDataCompression
	{
		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterNone = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterRLE = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterGroup3 = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterPackbits = 3,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterGroup4 = 4,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterLZW = 5,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visRasterModifiedHuffman = 6
	}
}