﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.PpSlideSizeType"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpSlideSizeType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSizeOnScreen = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSizeLetterPaper = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSizeA4Paper = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSize35MM = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSizeOverhead = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSizeBanner = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		 ppSlideSizeCustom = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeLedgerPaper = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeA3Paper = 9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeB4ISOPaper = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeB5ISOPaper = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeB4JISPaper = 12,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeB5JISPaper = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 ppSlideSizeHagakiCard = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppSlideSizeOnScreen16x9 = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("PowerPoint", 12,14,15,16)]
		 ppSlideSizeOnScreen16x10 = 16
	}
}