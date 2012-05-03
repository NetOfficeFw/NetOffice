using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdTaskPanes
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdTaskPaneFormatting = 0,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdTaskPaneRevealFormatting = 1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdTaskPaneMailMerge = 2,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdTaskPaneTranslate = 3,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdTaskPaneSearch = 4,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneXMLStructure = 5,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneDocumentProtection = 6,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneDocumentActions = 7,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneSharedWorkspace = 8,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneHelp = 9,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneResearch = 10,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneFaxService = 11,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneXMLDocument = 12,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdTaskPaneDocumentUpdates = 13,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdTaskPaneSignature = 14,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdTaskPaneStyleInspector = 15,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdTaskPaneDocumentManagement = 16,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdTaskPaneApplyStyles = 17,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdTaskPaneNav = 18,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdTaskPaneSelection = 19
	}
}