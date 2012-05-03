using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdMoveFromTextMark
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkHidden = 0,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkDoubleStrikeThrough = 1,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkStrikeThrough = 2,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkCaret = 3,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkPound = 4,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkNone = 5,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkBold = 6,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkItalic = 7,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkUnderline = 8,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkDoubleUnderline = 9,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMoveFromTextMarkColorOnly = 10
	}
}