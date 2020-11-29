using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoTextStrike"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoTextStrike
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoStrikeMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoNoStrike = 0,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoSingleStrike = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoDoubleStrike = 2
	}
}