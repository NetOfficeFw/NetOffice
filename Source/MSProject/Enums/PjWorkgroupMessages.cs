using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjWorkgroupMessages
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjDefaultMessages = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjNoWorkgroupMessages = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjWorkgroupViaEmail = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjWorkgroupViaWeb = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjWorkgroupViaEmailAndWeb = 4
	}
}