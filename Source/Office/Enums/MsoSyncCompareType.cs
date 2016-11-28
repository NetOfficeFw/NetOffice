using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Office", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoSyncCompareType
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15,16)]
		 msoSyncCompareAndMerge = 0,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15,16)]
		 msoSyncCompareSideBySide = 1
	}
}