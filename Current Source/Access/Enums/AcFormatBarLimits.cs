using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcFormatBarLimits
	{
		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acAutomatic = 0,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acNumber = 1,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acPercent = 2
	}
}