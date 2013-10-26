using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838491.aspx </remarks>
	[SupportByVersionAttribute("Word", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdOMathVertAlignType
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdOMathVertAlignCenter = 0,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdOMathVertAlignTop = 1,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdOMathVertAlignBottom = 2
	}
}