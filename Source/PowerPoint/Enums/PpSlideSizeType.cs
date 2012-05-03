using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpSlideSizeType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSizeOnScreen = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSizeLetterPaper = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSizeA4Paper = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSize35MM = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSizeOverhead = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSizeBanner = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppSlideSizeCustom = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeLedgerPaper = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeA3Paper = 9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeB4ISOPaper = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeB5ISOPaper = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeB4JISPaper = 12,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeB5JISPaper = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppSlideSizeHagakiCard = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppSlideSizeOnScreen16x9 = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppSlideSizeOnScreen16x10 = 16
	}
}