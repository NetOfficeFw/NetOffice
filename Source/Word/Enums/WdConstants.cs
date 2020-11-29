using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdConstants"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdConstants
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9999999</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdUndefined = 9999999,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9999998</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdToggle = 9999998,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1073741823</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdForward = 1073741823,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1073741823</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdBackward = -1073741823,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdAutoPosition = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdFirst = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1297307460</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCreatorCode = 1297307460
	}
}