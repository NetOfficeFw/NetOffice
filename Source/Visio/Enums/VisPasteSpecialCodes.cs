using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisPasteSpecialCodes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteText = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteBitmap = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteMetafile = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteOEMText = 7,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteDIB = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteEMF = 14,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteOLEObject = 65536,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65537</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteRichText = 65537,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65538</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteHyperlink = 65538,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65539</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteURL = 65539,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65540</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteVisioShapes = 65540,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65541</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteVisioMasters = 65541,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65542</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteVisioText = 65542,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65543</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteVisioIcon = 65543,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65544</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteInk = 65544,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65545</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteVisioShapesXML = 65545,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>65546</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visPasteVisioMastersXML = 65546,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>65548</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visPasteVisioShapesWithoutDataLinks = 65548
	}
}