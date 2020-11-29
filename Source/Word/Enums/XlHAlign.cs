using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.xlhalign"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlHAlign
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignCenterAcrossSelection = 7,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4117</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignDistributed = -4117,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignFill = 5,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignGeneral = 1,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4130</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignJustify = -4130,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4131</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignLeft = -4131,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4152</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHAlignRight = -4152
	}
}