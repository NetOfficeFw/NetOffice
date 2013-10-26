using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192751.aspx </remarks>
	[SupportByVersionAttribute("Word", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdSmartTagControlType
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlSmartTag = 1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlLink = 2,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlHelp = 3,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlHelpURL = 4,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlSeparator = 5,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlButton = 6,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlLabel = 7,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlImage = 8,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlCheckbox = 9,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlTextbox = 10,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlListbox = 11,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlCombo = 12,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlActiveX = 13,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlDocumentFragment = 14,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlDocumentFragmentURL = 15,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdControlRadioGroup = 16
	}
}