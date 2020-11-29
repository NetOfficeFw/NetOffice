using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoFileValidationMode"/> </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoFileValidationMode
	{
		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoFileValidationDefault = 0,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoFileValidationSkip = 1
	}
}