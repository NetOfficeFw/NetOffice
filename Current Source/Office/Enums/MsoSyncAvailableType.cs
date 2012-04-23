using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoSyncAvailableType
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoSyncAvailableNone = 0,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoSyncAvailableOffline = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoSyncAvailableAnywhere = 2
	}
}