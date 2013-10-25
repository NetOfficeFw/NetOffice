using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836278.aspx </remarks>
	[SupportByVersionAttribute("Access", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcExportXMLOtherFlags
	{
		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acEmbedSchema = 1,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acExcludePrimaryKeyAndIndexes = 2,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acRunFromServer = 4,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acLiveReportSource = 8,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acPersistReportML = 16,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acExportAllTableAndFieldProperties = 32
	}
}