using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlChartOrientation
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 xlVertical = -4166
	}
}