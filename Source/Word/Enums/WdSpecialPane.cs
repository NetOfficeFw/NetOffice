using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdSpecialPane
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneNone = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPanePrimaryHeader = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneFirstPageHeader = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneEvenPagesHeader = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPanePrimaryFooter = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneFirstPageFooter = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneEvenPagesFooter = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneFootnotes = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneEndnotes = 8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneFootnoteContinuationNotice = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneFootnoteContinuationSeparator = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneFootnoteSeparator = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneEndnoteContinuationNotice = 12,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneEndnoteContinuationSeparator = 13,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneEndnoteSeparator = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneComments = 15,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneCurrentPageHeader = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPaneCurrentPageFooter = 17,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdPaneRevisions = 18,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdPaneRevisionsHoriz = 19,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdPaneRevisionsVert = 20
	}
}