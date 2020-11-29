using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoPickerField"/> </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoPickerField
	{
		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoPickerFieldUnknown = 0,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoPickerFieldDateTime = 1,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoPickerFieldNumber = 2,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoPickerFieldText = 3,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoPickerFieldUser = 4,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoPickerFieldMax = 5
	}
}