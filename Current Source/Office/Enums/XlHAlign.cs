using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlHAlign
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignCenterAcrossSelection = 7,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4117</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignDistributed = -4117,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignFill = 5,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignGeneral = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4130</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignJustify = -4130,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignLeft = -4131,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlHAlignRight = -4152
	}
}