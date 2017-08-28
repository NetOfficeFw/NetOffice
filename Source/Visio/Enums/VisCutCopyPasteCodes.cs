using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisCutCopyPasteCodes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCopyPasteNormal = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCopyPasteNoTranslate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Visio", 14,15,16)]
		 visCopyPasteCenter = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Visio", 14,15,16)]
		 visCopyPasteNoHealConnectors = 4,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Visio", 14,15,16)]
		 visCopyPasteNoContainerMembers = 8,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Visio", 14,15,16)]
		 visCopyPasteNoAssociatedCallouts = 16,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Visio", 14,15,16)]
		 visCopyPasteDontAddToContainers = 32,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Visio", 14,15,16)]
		 visCopyPasteNoCascade = 64
	}
}