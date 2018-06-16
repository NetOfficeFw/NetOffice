using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPControls 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("914934A2-5A91-11CF-8700-00AA0060263B")]
	public interface PPControls : Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.PPControl this[object index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPPushButton AddPushButton(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPToggleButton AddToggleButton(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPBitmapButton AddBitmapButton(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPListBox AddListBox(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPCheckBox AddCheckBox(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPRadioCluster AddRadioCluster(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPStaticText AddStaticText(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="verticalScrollBar">optional object verticalScrollBar</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPEditText AddEditText(Single left, Single top, Single width, Single height, object verticalScrollBar);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPEditText AddEditText(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPIcon AddIcon(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPBitmap AddBitmap(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPSpinner AddSpinner(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="style">NetOffice.PowerPointApi.Enums.PpScrollBarStyle style</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPScrollBar AddScrollBar(NetOffice.PowerPointApi.Enums.PpScrollBarStyle style, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPGroupBox AddGroupBox(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDropDown AddDropDown(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPDropDownEdit AddDropDownEdit(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPSlideMiniature AddMiniature(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPFrame AddFrame(Single left, Single top, Single width, Single height);

		#endregion
	}
}
