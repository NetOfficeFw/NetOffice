using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoPickerField
	{
		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 msoPickerFieldUnknown = 0,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 msoPickerFieldDateTime = 1,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 msoPickerFieldNumber = 2,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 msoPickerFieldText = 3,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 msoPickerFieldUser = 4,

		 /// <summary>
		 /// SupportByVersion Office 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 14,15)]
		 msoPickerFieldMax = 5
	}
}