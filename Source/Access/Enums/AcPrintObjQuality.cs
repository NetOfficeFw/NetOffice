using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcPrintObjQuality"/> </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcPrintObjQuality
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPQDraft = -1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPQLow = -2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPQMedium = -3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPQHigh = -4
	}
}