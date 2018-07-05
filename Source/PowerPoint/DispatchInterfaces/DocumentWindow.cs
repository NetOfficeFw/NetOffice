using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface DocumentWindow
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("91493457-5A91-11CF-8700-00AA0060263B")]
	public interface DocumentWindow : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Selection Selection { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.View View { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Presentation { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpViewType ViewType { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState BlackAndWhite { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Active { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpWindowState WindowState { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Caption { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Left { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Top { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Width { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Height { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 HWND { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Pane ActivePane { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Panes Panes { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 SplitVertical { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 SplitHorizontal { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FitToPage();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		/// <param name="toLeft">optional Int32 ToLeft = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void LargeScroll(object down, object up, object toRight, object toLeft);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void LargeScroll();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void LargeScroll(object down);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void LargeScroll(object down, object up);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void LargeScroll(object down, object up, object toRight);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		/// <param name="toLeft">optional Int32 ToLeft = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SmallScroll(object down, object up, object toRight, object toLeft);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SmallScroll();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SmallScroll(object down);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SmallScroll(object down, object up);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SmallScroll(object down, object up, object toRight);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.DocumentWindow NewWindow();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		object RangeFromPoint(Int32 x, Int32 y);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="points">Single points</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 PointsToScreenPixelsX(Single points);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="points">Single points</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 PointsToScreenPixelsY(Single points);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="start">optional NetOffice.OfficeApi.Enums.MsoTriState Start = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void ScrollIntoView(Single left, Single top, Single width, Single height, object start);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void ScrollIntoView(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool IsSectionExpanded(Int32 sectionIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="expand">bool expand</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void ExpandSection(Int32 sectionIndex, bool expand);

		#endregion
	}
}
