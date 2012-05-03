using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdDeletedTextMark
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDeletedTextMarkHidden = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDeletedTextMarkStrikeThrough = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDeletedTextMarkCaret = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDeletedTextMarkPound = 3,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDeletedTextMarkNone = 4,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDeletedTextMarkBold = 5,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDeletedTextMarkItalic = 6,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDeletedTextMarkUnderline = 7,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDeletedTextMarkDoubleUnderline = 8,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDeletedTextMarkColorOnly = 9,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDeletedTextMarkDoubleStrikeThrough = 10
	}
}