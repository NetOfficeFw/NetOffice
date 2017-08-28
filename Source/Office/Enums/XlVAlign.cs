using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlVAlign
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlVAlignBottom = -4107,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlVAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4117</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlVAlignDistributed = -4117,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4130</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlVAlignJustify = -4130,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 xlVAlignTop = -4160
	}
}