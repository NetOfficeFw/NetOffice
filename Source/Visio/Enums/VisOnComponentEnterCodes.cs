using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisOnComponentEnterCodes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visComponentStateModal = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visModalDeferEvents = 65536,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visModalNoBeforeAfter = 131072,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visModalDontBlockMessages = 262144,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visModalDisableVisiosFrame = 524288
	}
}