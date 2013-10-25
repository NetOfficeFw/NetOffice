using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195243.aspx </remarks>
	[SupportByVersionAttribute("Access", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcPrintObjQuality
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acPRPQDraft = -1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acPRPQLow = -2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acPRPQMedium = -3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acPRPQHigh = -4
	}
}