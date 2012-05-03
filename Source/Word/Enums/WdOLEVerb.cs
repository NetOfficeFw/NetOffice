using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdOLEVerb
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbPrimary = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbShow = -1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbOpen = -2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbHide = -3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbUIActivate = -4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbInPlaceActivate = -5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOLEVerbDiscardUndoState = -6
	}
}