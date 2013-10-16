using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoMixedType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoIntegerMixed = 32768,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoSingleMixed = -2147483648
	}
}