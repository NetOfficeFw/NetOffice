using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdRectangleType
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTextRectangle = 0,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdShapeRectangle = 1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdMarkupRectangle = 2,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdMarkupRectangleButton = 3,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdPageBorderRectangle = 4,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdLineBetweenColumnRectangle = 5,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdSelection = 6,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdSystem = 7,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMarkupRectangleArea = 8,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdReadingModeNavigation = 9,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMarkupRectangleMoveMatch = 10,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdReadingModePanningArea = 11,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdMailNavArea = 12,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDocumentControlRectangle = 13
	}
}