using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoSyncCompareType
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoSyncCompareAndMerge = 0,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoSyncCompareSideBySide = 1
	}
}