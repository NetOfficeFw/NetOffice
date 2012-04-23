using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoIodGroup
	{
		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoIodGroupPIAs = 0,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoIodGroupVSTOR35Mgd = 1,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoIodGroupVSTOR40Mgd = 2
	}
}