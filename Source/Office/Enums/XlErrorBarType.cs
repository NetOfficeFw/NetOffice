using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlErrorBarType
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4114</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlErrorBarTypeCustom = -4114,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlErrorBarTypeFixedValue = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlErrorBarTypePercent = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4155</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlErrorBarTypeStDev = -4155,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlErrorBarTypeStError = 4
	}
}