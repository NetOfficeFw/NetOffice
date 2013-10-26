using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198108.aspx </remarks>
	[SupportByVersionAttribute("Word", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdLineType
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdTextLine = 0,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdTableRow = 1
	}
}