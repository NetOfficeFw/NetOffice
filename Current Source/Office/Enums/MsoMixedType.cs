using System;
using LateBindingApi.Core;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoMixedType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoIntegerMixed = 32768,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoSingleMixed = -2147483648
	}
}