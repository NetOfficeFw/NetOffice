using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisInsertObjArgs
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visInsertLink = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visInsertIcon = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visInsertDontShow = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visInsertAsControl = 8192,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visInsertAsEmbed = 16384,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visInsertNoDesignModeTransition = 256
	}
}