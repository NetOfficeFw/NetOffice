using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoTextDirection"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoTextDirection
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoTextDirectionMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoTextDirectionLeftToRight = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoTextDirectionRightToLeft = 2
	}
}