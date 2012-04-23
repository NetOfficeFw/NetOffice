using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdConstants
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9999999</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdUndefined = 9999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9999998</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdToggle = 9999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1073741823</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdForward = 1073741823,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1073741823</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdBackward = -1073741823,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdAutoPosition = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFirst = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1297307460</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdCreatorCode = 1297307460
	}
}