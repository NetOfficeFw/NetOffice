﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdInlineShapeType"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdInlineShapeType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeEmbeddedOLEObject = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeLinkedOLEObject = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapePicture = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeLinkedPicture = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeOLEControlObject = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeHorizontalLine = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapePictureHorizontalLine = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeLinkedPictureHorizontalLine = 8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapePictureBullet = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeScriptAnchor = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdInlineShapeOWSAnchor = 11,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdInlineShapeChart = 12,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdInlineShapeDiagram = 13,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdInlineShapeLockedCanvas = 14,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdInlineShapeSmartArt = 15,

		 /// <summary>
		 /// SupportByVersion Word 15,16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Word", 15, 16)]
		 wdInlineShapeWebVideo = 16
	}
}