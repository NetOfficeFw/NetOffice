using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Access", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcExportQuality
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acExportQualityPrint = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acExportQualityScreen = 1
	}
}