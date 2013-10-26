using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765180(v=office.14).aspx </remarks>
	[SupportByVersionAttribute("Visio", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisRasterExportColorFormat
	{
		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRasterBiLevel = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRaster16Color = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRaster256Color = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRaster24Bit = 3,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRasterRGB = 4,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRasterYCC = 5,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRasterGrayScale = 6,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRasterCMYK = 7,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRaster16ColorGrayScale = 8,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRaster256ColorGrayScale = 9,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRaster16Bit = 10
	}
}