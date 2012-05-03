using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoSyncVersionType
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoSyncVersionLastViewed = 0,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoSyncVersionServer = 1
	}
}