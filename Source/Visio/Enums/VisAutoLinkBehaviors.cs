using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisAutoLinkBehaviors
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkSelectedShapesOnly = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkGenericProgressBar = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkNoApplyDataGraphic = 4,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkReplaceExistingLinks = 8,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkDontReplaceExistingLinks = 16,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkNullMatchesNoFormula = 32,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoLinkIncludeHiddenProps = 64
	}
}