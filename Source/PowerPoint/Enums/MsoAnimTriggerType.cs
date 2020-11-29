using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	/// <summary>
	/// The action that triggers the animation effect. The default value is <see cref="msoAnimTriggerOnPageClick"/>.
	/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
	/// </summary>
	///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.MsoAnimTriggerType"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAnimTriggerType
	{
		/// <summary>
		/// Mixed actions.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>-1</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimTriggerMixed = -1,

		/// <summary>
		/// No action associated as the trigger.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>0</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimTriggerNone = 0,

		/// <summary>
		/// When a page is clicked.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>1</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimTriggerOnPageClick = 1,

		/// <summary>
		/// When the Previous button is clicked.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>2</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimTriggerWithPrevious = 2,

		/// <summary>
		/// After the Previous button is clicked.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>3</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimTriggerAfterPrevious = 3,

		/// <summary>
		/// When a shape is clicked.
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>4</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimTriggerOnShapeClick = 4,

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks>5</remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		msoAnimTriggerOnMediaBookmark = 5
	}
}