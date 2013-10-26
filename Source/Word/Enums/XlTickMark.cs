using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822691.aspx </remarks>
	[SupportByVersionAttribute("Word", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlTickMark
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickMarkCross = 4,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickMarkInside = 2,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>-4142</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickMarkNone = -4142,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 xlTickMarkOutside = 3
	}
}