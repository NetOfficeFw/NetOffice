using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpViewType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewSlide = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewSlideMaster = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewNotesPage = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewHandoutMaster = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewNotesMaster = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewOutline = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewSlideSorter = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewTitleMaster = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppViewNormal = 9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppViewPrintPreview = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppViewThumbnails = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppViewMasterThumbnails = 12
	}
}