using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlScaleType
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4132</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlScaleLinear = -4132,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4133</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlScaleLogarithmic = -4133
	}
}