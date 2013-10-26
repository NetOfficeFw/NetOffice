using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834917.aspx </remarks>
	[SupportByVersionAttribute("Word", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlTrendlineType
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlExponential = 5,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4132</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlLinear = -4132,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4133</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlLogarithmic = -4133,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlMovingAvg = 6,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlPolynomial = 3,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlPower = 4
	}
}