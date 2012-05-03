using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdXMLSelectionChangeReason
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLSelectionChangeReasonMove = 0,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLSelectionChangeReasonInsert = 1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdXMLSelectionChangeReasonDelete = 2
	}
}