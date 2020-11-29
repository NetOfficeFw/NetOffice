﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdRevisionType"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdRevisionType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoRevision = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionInsert = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionDelete = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionProperty = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionParagraphNumber = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionDisplayField = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionReconcile = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionConflict = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionStyle = 8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdRevisionReplace = 9,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdRevisionParagraphProperty = 10,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdRevisionTableProperty = 11,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdRevisionSectionProperty = 12,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdRevisionStyleDefinition = 13,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdRevisionMovedFrom = 14,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdRevisionMovedTo = 15,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdRevisionCellInsertion = 16,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdRevisionCellDeletion = 17,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdRevisionCellMerge = 18,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdRevisionCellSplit = 19,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdRevisionConflictInsert = 20,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdRevisionConflictDelete = 21
	}
}