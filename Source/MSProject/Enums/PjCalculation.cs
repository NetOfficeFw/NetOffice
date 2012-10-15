using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjCalculation
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjAutomatic = -1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjManual = 0
	}
}